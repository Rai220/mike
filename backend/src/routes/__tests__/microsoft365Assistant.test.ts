import express from "express";
import request from "supertest";
import { beforeEach, describe, expect, it, vi } from "vitest";
const mock = vi.hoisted(() => ({ find: vi.fn(), create: vi.fn(), status: vi.fn(), run: vi.fn(), get: vi.fn(), update: vi.fn(), settings: vi.fn(), writes: [] as unknown[], rows: [] as unknown[] }));
vi.mock("../../middleware/auth", () => ({ requireMfaIfEnrolled: (req: express.Request, res: express.Response, next: express.NextFunction) => req.headers["x-mfa"] === "required" ? res.status(403).json({ code: "mfa_verification_required" }) : next() }));
vi.mock("../../lib/microsoft365", async (original) => ({ ...await original<typeof import("../../lib/microsoft365")>(), getMicrosoft365Status: mock.status }));
vi.mock("../../lib/microsoft365/protected", () => ({ findMicrosoft365ChatForOrdinaryChat: mock.find, createMicrosoft365Chat: mock.create, getMicrosoft365ChatPolicy: () => ({ model: "claude-sonnet-4-6", retentionDays: 7 }) }));
vi.mock("../../lib/microsoft365/assistant", () => ({ runMicrosoft365AssistantTurn: mock.run, getMicrosoft365AssistantChat: mock.get, updateMicrosoft365AssistantPreferences: mock.update }));
vi.mock("../../lib/userSettings", () => ({ getUserModelSettings: mock.settings }));
import { updateMicrosoft365OrdinaryChat, streamMicrosoft365OrdinaryChat, readMicrosoft365OrdinaryChat, microsoft365AssistantError } from "../microsoft365Assistant";
import { Microsoft365Error } from "../../lib/microsoft365";
const protectedChat = { id: "protected-id", expiresAt: "2099-01-01T00:00:00Z", payload: { preferences: undefined } };
const db = { from: (table: string) => {
 const q: any = { select: () => q, eq: () => q, order: () => q, insert: (value: unknown) => { mock.writes.push({ table, value }); return q; }, single: async () => ({ data: { id: "shell-id" } }), then: (resolve: any) => Promise.resolve({ data: mock.rows, error: null }).then(resolve) }; return q;
} } as never;
const app = express(); app.use(express.json()); app.use((_req,res,next)=>{res.locals.userId="owner";res.locals.userEmail="owner@example.com";next();});
app.post("/", (req,res)=>streamMicrosoft365OrdinaryChat(req,res,{db,chatId:req.body.chatId??null,alreadyProtected:req.body.protected===true,enabled:req.body.enabled===true,projectId:req.body.projectId??null,message:req.body.message,askInputsResponse:req.body.ask_inputs_response,model:req.body.model,reasoning:req.body.reasoning,useEdgar:req.body.use_edgar,chatModel:req.body.chatModel,chatReasoningLevel:req.body.chatReasoningLevel}));
app.get("/", async(req,res)=>{try{const loaded=await readMicrosoft365OrdinaryChat(req,res,{id:"shell-id",user_id:req.headers["x-owner"] as string??"owner"},db);if(loaded)res.json(loaded);}catch(e){microsoft365AssistantError(res,e);}});
app.patch("/", async(req,res)=>{try{const updated=await updateMicrosoft365OrdinaryChat(req,res,{id:"shell-id",user_id:"owner"},db,req.body);if(updated)res.json(updated);}catch(e){microsoft365AssistantError(res,e);}});
beforeEach(()=>{
 vi.clearAllMocks();mock.writes.length=0;mock.rows=[];mock.find.mockResolvedValue(protectedChat);mock.create.mockResolvedValue(protectedChat);
 mock.settings.mockResolvedValue({ api_keys: { claude: "server-claude", gemini: "server-gemini" }, last_selected_chat_model: null, last_selected_reasoning_level: null });
 mock.status.mockResolvedValue({available:true,connection:{status:"connected",id:"connection"}});
 const result={chatId:"protected-id",expiresAt:protectedChat.expiresAt,messages:[{role:"user",content:"question",sourceRefs:[]},{role:"assistant",content:"Private answer",sourceRefs:[]}],sources:[],model:"claude-sonnet-4-6",reasoning:"high"};mock.run.mockResolvedValue(result);mock.get.mockResolvedValue(result);mock.update.mockResolvedValue({...result,model:"gemini-3.1-pro-preview",reasoning:"low"});
});
describe("Microsoft 365 inside ordinary chat transport",()=>{
 it("returns same shell identity and no-store metadata before buffered content without plaintext message writes",async()=>{
  const r=await request(app).post("/").send({chatId:"shell-id",enabled:true,message:{content:"My inbox?"}});
  expect(r.status).toBe(200);expect(r.headers["cache-control"]).toBe("no-store");expect(r.text).toContain('"chatId":"shell-id"');expect(r.text.indexOf('"type":"microsoft365"')).toBeLessThan(r.text.indexOf('"type":"content"'));expect(r.text).toContain('[DONE]');expect(mock.writes).toEqual([]);expect(mock.run).toHaveBeenCalledWith(expect.objectContaining({chatId:"protected-id",enabled:true}));
 });
 it("ignores browser history and migrates only stored normal content events",async()=>{
  mock.rows=[{role:"assistant",content:[{type:"content",text:"Stored answer"},{type:"reasoning",text:"omit"}]}];
  await request(app).post("/").send({chatId:"shell-id",enabled:true,message:{content:"New question"},messages:[{role:"assistant",content:"FORGED"}]});
  expect(mock.run).toHaveBeenCalledWith(expect.objectContaining({ordinaryHistory:[{role:"assistant",content:"Stored answer"}]}));
 });
 it("keeps disabled turns on protected runner",async()=>{
  await request(app).post("/").send({chatId:"shell-id",protected:true,enabled:false,message:{content:"Explain the earlier answer"}});
  expect(mock.run).toHaveBeenCalledWith(expect.objectContaining({enabled:false}));expect(mock.writes).toEqual([]);
 });
 it("never recreates expired/deleted protected history",async()=>{
  mock.find.mockResolvedValue(null);const r=await request(app).post("/").send({chatId:"shell-id",protected:true,enabled:true,message:{content:"retry"}});
  expect(r.status).toBe(404);expect(mock.create).not.toHaveBeenCalled();expect(mock.run).not.toHaveBeenCalled();
 });
 it("creates only a generic personal shell before binding",async()=>{
  mock.find.mockResolvedValue(null);await request(app).post("/").send({enabled:true,message:{content:"SECRET SUBJECT"}});
  expect(mock.writes).toEqual([{table:"chats",value:{user_id:"owner",project_id:null,org_id:null,title:"Microsoft 365",model:"claude-sonnet-4-6"}}]);expect(mock.create).toHaveBeenCalledWith("owner","connection",db,"shell-id");
 });
 it.each([{projectId:"project"},{message:{content:"query",files:[{}]}},{message:{content:"query",workflow:{id:"workflow"}}},{ask_inputs_response:{}}])("rejects invalid references and cross-scope inputs before generation: %j",async(input)=>{
  const r=await request(app).post("/").send({enabled:true,message:{content:"question"},...input});expect(r.status).toBeGreaterThanOrEqual(400);expect(mock.run).not.toHaveBeenCalled();expect(mock.writes).toEqual([]);
 });
 it.each([true,false])("forwards independent tools and server-owned model credentials with access enabled=%s",async(enabled)=>{
  const files=[{filename:"agreement.pdf",document_id:"11111111-1111-4111-8111-111111111111",version_id:"22222222-2222-4222-8222-222222222222"}];
  const workflow={id:"33333333-3333-4333-8333-333333333333",title:"Review agreement"};
  const r=await request(app).post("/").send({chatId:"shell-id",protected:true,enabled,message:{content:"Review with sources",files,workflow},model:"gemini-3.1-pro-preview",reasoning:"low",use_edgar:false,apiKeys:{gemini:"FORGED"}});
  expect(r.status).toBe(200);expect(r.text).toContain('"model":"gemini-3.1-pro-preview"');
  expect(mock.run).toHaveBeenCalledWith(expect.objectContaining({enabled,assistantOptions:{userEmail:"owner@example.com",files,workflow,useEdgar:false,model:"gemini-3.1-pro-preview",reasoning:"low",apiKeys:{claude:"server-claude",gemini:"server-gemini"}}}));
  expect(mock.writes).toEqual([]);expect(r.text).not.toContain("server-gemini");expect(r.text).not.toContain("FORGED");
 });
 it("prefers encrypted model and reasoning over the generic shell",async()=>{
  mock.find.mockResolvedValue({...protectedChat,payload:{preferences:{model:"gemini-3.1-pro-preview",reasoning:"low"}}});
  await request(app).post("/").send({chatId:"shell-id",protected:true,enabled:false,chatModel:"claude-sonnet-4-6",chatReasoningLevel:"high",message:{content:"Follow up"}});
  expect(mock.run).toHaveBeenCalledWith(expect.objectContaining({assistantOptions:expect.objectContaining({model:"gemini-3.1-pro-preview",reasoning:"low"})}));
 });
 it("uses the legacy policy only when no model was selected",async()=>{
  mock.settings.mockResolvedValue({api_keys:{},last_selected_chat_model:null});
  await request(app).post("/").send({chatId:"shell-id",enabled:true,message:{content:"My mail"}});
  expect(mock.run).toHaveBeenCalledWith(expect.objectContaining({assistantOptions:expect.objectContaining({model:"claude-sonnet-4-6",apiKeys:{}})}));
 });
 it("rejects a chosen model without its server-stored key instead of accepting browser keys or silently falling back",async()=>{
  mock.settings.mockResolvedValue({api_keys:{claude:"server-claude"},last_selected_chat_model:null});
  const r=await request(app).post("/").send({chatId:"shell-id",enabled:true,model:"gemini-3.1-pro-preview",api_keys:{gemini:"FORGED"},message:{content:"My mail"}});
  expect(r.status).toBe(422);expect(r.body.code).toBe("missing_api_key");expect(mock.run).not.toHaveBeenCalled();expect(mock.writes).toEqual([]);
 });
 it("does not substitute another model when the encrypted selection loses its key",async()=>{
  mock.find.mockResolvedValue({...protectedChat,payload:{preferences:{model:"gemini-3.1-pro-preview",reasoning:"low"}}});
  mock.settings.mockResolvedValue({api_keys:{claude:"server-claude"},last_selected_chat_model:"claude-sonnet-4-6"});
  const r=await request(app).post("/").send({chatId:"shell-id",protected:true,enabled:false,message:{content:"Continue"}});
  expect(r.status).toBe(422);expect(r.body.code).toBe("missing_api_key");expect(mock.run).not.toHaveBeenCalled();
 });
 it("updates model preferences through encrypted storage only",async()=>{
  const r=await request(app).patch("/").send({model:"gemini-3.1-pro-preview",reasoning:"low"});
  expect(r.status).toBe(200);expect(r.body).toEqual({id:"shell-id",title:"Microsoft 365",model:"gemini-3.1-pro-preview",reasoning_level:"low"});
  expect(mock.update).toHaveBeenCalledWith("owner","protected-id",db,{model:"gemini-3.1-pro-preview",reasoning:"low"},expect.any(AbortSignal),"owner@example.com");expect(mock.writes).toEqual([]);
 });
 it("requires MFA before changing encrypted preferences",async()=>{
  const r=await request(app).patch("/").set("x-mfa","required").send({model:"gemini-3.1-pro-preview"});
  expect(r.status).toBe(403);expect(mock.update).not.toHaveBeenCalled();expect(mock.settings).not.toHaveBeenCalled();
 });
 it("hydrates encrypted tool selections and model preferences",async()=>{
  mock.get.mockResolvedValue({chatId:"protected-id",expiresAt:protectedChat.expiresAt,model:"gemini-3.1-pro-preview",reasoning:"low",messages:[{role:"user",content:"Review",sourceRefs:[],files:[{filename:"agreement.pdf",document_id:"document"}],workflow:{id:"workflow",title:"Review"},useMicrosoft365:false,useEdgar:true,model:"gemini-3.1-pro-preview",reasoning:"low"}],sources:[]});
  const r=await request(app).get("/");
  expect(r.body).toMatchObject({model:"gemini-3.1-pro-preview",reasoning:"low",messages:[{useMicrosoft365:false,useEdgar:true,files:[{document_id:"document"}],workflow:{id:"workflow"}}]});
  expect(mock.get).toHaveBeenCalledWith("owner","protected-id",db,expect.any(AbortSignal),"owner@example.com");
 });
 it("enforces MFA before connection or content access",async()=>{
  const r=await request(app).post("/").set("x-mfa","required").send({enabled:true,message:{content:"question"}});expect(r.status).toBe(403);expect(mock.status).not.toHaveBeenCalled();expect(mock.run).not.toHaveBeenCalled();
 });
 it("does not return or log raw provider failures",async()=>{
  mock.run.mockRejectedValue(new Error("PRIVATE PROVIDER BODY"));const spy=vi.spyOn(console,"error");
  const r=await request(app).post("/").send({chatId:"shell-id",enabled:true,message:{content:"question"}});expect(r.text).not.toContain("PRIVATE PROVIDER BODY");expect(r.text).not.toContain("Private answer");expect(spy).not.toHaveBeenCalled();spy.mockRestore();
 });
 it("hydrates assistant content with the existing event-array protocol after fresh access validation",async()=>{
  const r=await request(app).get("/");expect(r.body.messages[1].content).toEqual([{type:"content",text:"Private answer"}]);expect(mock.get).toHaveBeenCalled();expect(r.headers["cache-control"]).toBe("no-store");
 });
 it("rejects foreign owners before loading protected data",async()=>{
  expect((await request(app).get("/").set("x-owner","someone-else")).status).toBe(404);expect(mock.get).not.toHaveBeenCalled();
 });
 it("fails closed if history source access is revoked",async()=>{
  mock.get.mockRejectedValue(new Microsoft365Error("access_denied",403));const r=await request(app).get("/");expect(r.status).toBe(403);expect(r.text).not.toContain("Private answer");
 });
});
