import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";

// ── Mocks ─────────────────────────────────────────────────────────────────────

const mockInitWithMiddleware = vi.fn().mockReturnValue({ api: vi.fn() });
const mockClientSecretCredential = vi.fn();
const mockTokenCredentialAuthenticationProvider = vi.fn();

vi.mock("@azure/identity", () => ({
  ClientSecretCredential: mockClientSecretCredential,
}));

vi.mock("@microsoft/microsoft-graph-client", () => ({
  Client: {
    initWithMiddleware: mockInitWithMiddleware,
  },
}));

vi.mock(
  "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js",
  () => ({
    TokenCredentialAuthenticationProvider: mockTokenCredentialAuthenticationProvider,
  })
);

// ── Helpers ───────────────────────────────────────────────────────────────────

const VALID_ENV = {
  TENANT_ID: "test-tenant",
  CLIENT_ID: "test-client",
  CLIENT_SECRET: "test-secret",
};

function setEnv(vars: Partial<typeof VALID_ENV>) {
  for (const [key, value] of Object.entries(vars)) {
    process.env[key] = value;
  }
}

function clearEnv() {
  delete process.env.TENANT_ID;
  delete process.env.CLIENT_ID;
  delete process.env.CLIENT_SECRET;
}

// ── Tests ─────────────────────────────────────────────────────────────────────

describe("createGraphClient", () => {
  beforeEach(() => {
    clearEnv();
    vi.clearAllMocks();
  });

  afterEach(() => {
    clearEnv();
  });

  describe("missing environment variables", () => {
    it("throws when all env vars are missing", async () => {
      const { createGraphClient } = await import("../graph-client.js");
      expect(() => createGraphClient()).toThrow(
        "Missing required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET"
      );
    });

    it("throws when TENANT_ID is missing", async () => {
      setEnv({ CLIENT_ID: VALID_ENV.CLIENT_ID, CLIENT_SECRET: VALID_ENV.CLIENT_SECRET });
      const { createGraphClient } = await import("../graph-client.js");
      expect(() => createGraphClient()).toThrow(
        "Missing required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET"
      );
    });

    it("throws when CLIENT_ID is missing", async () => {
      setEnv({ TENANT_ID: VALID_ENV.TENANT_ID, CLIENT_SECRET: VALID_ENV.CLIENT_SECRET });
      const { createGraphClient } = await import("../graph-client.js");
      expect(() => createGraphClient()).toThrow(
        "Missing required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET"
      );
    });

    it("throws when CLIENT_SECRET is missing", async () => {
      setEnv({ TENANT_ID: VALID_ENV.TENANT_ID, CLIENT_ID: VALID_ENV.CLIENT_ID });
      const { createGraphClient } = await import("../graph-client.js");
      expect(() => createGraphClient()).toThrow(
        "Missing required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET"
      );
    });
  });

  describe("successful client creation", () => {
    beforeEach(() => {
      setEnv(VALID_ENV);
    });

    it("creates ClientSecretCredential with correct args", async () => {
      const { createGraphClient } = await import("../graph-client.js");
      createGraphClient();
      expect(mockClientSecretCredential).toHaveBeenCalledOnce();
      expect(mockClientSecretCredential).toHaveBeenCalledWith(
        VALID_ENV.TENANT_ID,
        VALID_ENV.CLIENT_ID,
        VALID_ENV.CLIENT_SECRET
      );
    });

    it("creates TokenCredentialAuthenticationProvider with Graph default scope", async () => {
      const fakeCredential = {};
      mockClientSecretCredential.mockImplementationOnce(function() { return fakeCredential; });

      const { createGraphClient } = await import("../graph-client.js");
      createGraphClient();

      expect(mockTokenCredentialAuthenticationProvider).toHaveBeenCalledOnce();
      expect(mockTokenCredentialAuthenticationProvider).toHaveBeenCalledWith(fakeCredential, {
        scopes: ["https://graph.microsoft.com/.default"],
      });
    });

    it("calls Client.initWithMiddleware with the auth provider and returns its result", async () => {
      const fakeAuthProvider = {};
      mockTokenCredentialAuthenticationProvider.mockImplementationOnce(function() { return fakeAuthProvider; });
      const fakeClient = { api: vi.fn() };
      mockInitWithMiddleware.mockReturnValueOnce(fakeClient);

      const { createGraphClient } = await import("../graph-client.js");
      const result = createGraphClient();

      expect(mockInitWithMiddleware).toHaveBeenCalledOnce();
      expect(mockInitWithMiddleware).toHaveBeenCalledWith({ authProvider: fakeAuthProvider });
      expect(result).toBe(fakeClient);
    });
  });
});
