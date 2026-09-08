import { beforeEach, describe, expect, it, vi } from "vitest";
import { maintainMicrosoft365 } from "./maintenance";
import {
  getMicrosoft365Config,
  rotateMicrosoft365ConnectionEncryption,
} from "./index";
import {
  cleanupMicrosoft365Chats,
  rotateMicrosoft365ChatEncryption,
} from "./protected";
import { enqueueDbJob } from "../dbq/enqueue";
vi.mock("./index", () => ({
  getMicrosoft365Config: vi.fn(),
  rotateMicrosoft365ConnectionEncryption: vi.fn(),
}));
vi.mock("./protected", () => ({
  cleanupMicrosoft365Chats: vi.fn(),
  rotateMicrosoft365ChatEncryption: vi.fn(),
}));
vi.mock("../dbq/enqueue", () => ({ enqueueDbJob: vi.fn() }));
const db = {} as Parameters<typeof maintainMicrosoft365>[0];
const job = { payload: {} } as Parameters<typeof maintainMicrosoft365>[1];
beforeEach(() => {
  vi.resetAllMocks();
  vi.mocked(getMicrosoft365Config).mockReturnValue(
    {} as NonNullable<ReturnType<typeof getMicrosoft365Config>>,
  );
  vi.mocked(cleanupMicrosoft365Chats).mockResolvedValue(3);
  vi.mocked(rotateMicrosoft365ConnectionEncryption).mockResolvedValue({
    examined: 2,
    rotated: 2,
    skipped: 0,
    nextCursor: null,
  });
  vi.mocked(rotateMicrosoft365ChatEncryption).mockResolvedValue({
    examined: 1,
    rotated: 1,
    skipped: 0,
    nextCursor: null,
  });
});
describe("corporate retention and rotation jobs", () => {
  it("keeps erasing expired data when the connection feature is disabled", async () => {
    vi.mocked(getMicrosoft365Config).mockReturnValue(null);
    expect(await maintainMicrosoft365(db, job)).toEqual({ removed: 3 });
    expect(rotateMicrosoft365ConnectionEncryption).not.toHaveBeenCalled();
  });
  it("reports only counts and does not reschedule completed batches", async () => {
    expect(await maintainMicrosoft365(db, job)).toEqual({
      removed: 3,
      rotated: 3,
      skipped: 0,
    });
    expect(enqueueDbJob).not.toHaveBeenCalled();
  });
  it("durably continues bounded rotation without repeating finished table or serializing content", async () => {
    vi.mocked(rotateMicrosoft365ConnectionEncryption).mockResolvedValue({
      examined: 100,
      rotated: 98,
      skipped: 2,
      nextCursor: "next-uuid",
    });
    await maintainMicrosoft365(db, job);
    expect(enqueueDbJob).toHaveBeenCalledWith(
      db,
      expect.objectContaining({
        kind: "microsoft365.maintenance",
        payload: {
          connectionCursor: "next-uuid",
          chatCursor: null,
          connectionsDone: false,
          chatsDone: true,
        },
      }),
    );
    await maintainMicrosoft365(db, {
      payload: { connectionCursor: "next-uuid", chatsDone: true },
    } as typeof job);
    expect(rotateMicrosoft365ChatEncryption).toHaveBeenCalledTimes(1);
    expect(rotateMicrosoft365ConnectionEncryption).toHaveBeenLastCalledWith(
      db,
      "next-uuid",
    );
  });
  it("does not swallow cleanup failure or silently drop continuation jobs", async () => {
    vi.mocked(cleanupMicrosoft365Chats).mockRejectedValue(
      new Error("safe_failure"),
    );
    await expect(maintainMicrosoft365(db, job)).rejects.toThrow("safe_failure");
    expect(rotateMicrosoft365ConnectionEncryption).not.toHaveBeenCalled();
  });
});
