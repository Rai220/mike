# Documentation

## Run and deploy Mike

- [Local development](local-development.md) — Docker Compose, local services,
  registration, Ollama, and first-run setup
- [Manual and production deployment](deployment.md) — managed infrastructure,
  environment variables, database upgrades, and deployment safety
- [Troubleshooting](troubleshooting.md) — common local and production problems
- [Safe local testing](safe-local-testing.md) — disposable resources, synthetic
  documents, and secret handling

## Features and clients

- [CourtListener integration](courtlistener.md) — live US case-law tools and
  optional bulk data
- [SEC EDGAR integration](edgar.md) — public-company filing research tools
- [Microsoft Word add-in](../word-addin/README.md) — concise setup and command
  reference
- [Word add-in development and deployment](word-addin-development.md) — manual
  setup, sideloading, builds, storage behavior, testing, and troubleshooting
- [Tamper-evident exports](tamper-evident-exports.md) — document hashes and
  optional signed manifests

## Frontend

- [Design system](design-system.md) — color/typography/spacing tokens, the shared
  `components/ui` primitives, and the accessibility baseline

## Testing and CI

- [End-to-end tests in CI](e2e-ci.md)
- [Backend unit-test coverage](testing-coverage.md)
- [Frontend unit-test coverage](frontend-testing.md)
- [Mutation testing and the SSE load harness](test-depth.md)

## Historical design and investigation notes

These files preserve the context of completed work. They are not current setup
or architecture guidance.

- [Legal workflows design spec](superpowers/specs/2026-06-29-legal-workflows-design.md)
- [Word add-in assistant scroll-jump report](word-addin-chat-scroll-report.md)

## Microsoft 365 integration

- [Microsoft 365 integration plan](microsoft-365-integration-plan.md) — architecture
  original architecture and implementation decisions for Outlook, OneDrive, and SharePoint
- [Microsoft 365 connection check](microsoft365-connection-check.md) — local
  credential and delegated mail diagnostics; live canary evidence
- [Microsoft 365 connection setup](microsoft365-setup.md) — operator configuration
  and user flow for the protected mail/file web chat

- [Microsoft 365 protected web chat PRD](microsoft365-protected-chat-prd.md) — implemented scope, privacy boundaries, acceptance gaps.

- [Microsoft 365 verification](microsoft365-verification.md) — test evidence and live acceptance limits.

Contribution and disclosure policies live in [CONTRIBUTING.md](../CONTRIBUTING.md)
and [SECURITY.md](../SECURITY.md).

- [Microsoft 365 RnD deployment](microsoft365-rnd-deployment.md) — rollout, verification and activation status.
