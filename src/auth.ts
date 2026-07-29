import {
  LogLevel,
  PublicClientApplication,
  type AccountInfo,
  type Configuration,
  type ICachePlugin,
} from "@azure/msal-node";
import {
  DataProtectionScope,
  PersistenceCachePlugin,
  PersistenceCreator,
} from "@azure/msal-node-extensions";
import { mkdir } from "node:fs/promises";
import path from "node:path";
import type { MailboxConfig } from "./types.js";

const GRAPH_SCOPES = ["User.Read", "Mail.Read"];

export interface AuthEnvironment {
  clientId: string;
  tenantId: string;
}

export function readAuthEnvironment(): AuthEnvironment {
  const clientId = process.env.MS_CLIENT_ID?.trim();
  const tenantId = process.env.MS_TENANT_ID?.trim();
  if (clientId === undefined || clientId === "") {
    throw new Error(
      "MS_CLIENT_ID is missing. Copy .env.example to .env and add the Entra application client ID.",
    );
  }
  if (tenantId === undefined || tenantId === "") {
    throw new Error(
      "MS_TENANT_ID is missing. Copy .env.example to .env and add the Entra tenant ID.",
    );
  }
  return { clientId, tenantId };
}

async function createCachePlugin(
  cacheDirectory: string,
  mailboxId: string,
  clientId: string,
): Promise<ICachePlugin> {
  await mkdir(cacheDirectory, { recursive: true, mode: 0o700 });
  const cachePath = path.join(cacheDirectory, `${mailboxId}.cache`);
  const persistence = await PersistenceCreator.createPersistence({
    cachePath,
    dataProtectionScope: DataProtectionScope.CurrentUser,
    serviceName: `OutlookFirstResponseTracker-${clientId}`,
    accountName: mailboxId,
    usePlaintextFileOnLinux: false,
  });
  return new PersistenceCachePlugin(persistence);
}

async function createClient(
  mailbox: MailboxConfig,
  cacheDirectory: string,
  environment: AuthEnvironment,
): Promise<PublicClientApplication> {
  const cachePlugin = await createCachePlugin(
    cacheDirectory,
    mailbox.id,
    environment.clientId,
  );
  const configuration: Configuration = {
    auth: {
      clientId: environment.clientId,
      authority: `https://login.microsoftonline.com/${environment.tenantId}`,
    },
    cache: { cachePlugin },
    system: {
      loggerOptions: {
        logLevel: LogLevel.Warning,
        piiLoggingEnabled: false,
        loggerCallback: (_level, message, containsPii) => {
          if (!containsPii) {
            console.warn(`[Microsoft auth] ${message}`);
          }
        },
      },
    },
  };
  return new PublicClientApplication(configuration);
}

function findAccount(
  accounts: AccountInfo[],
  expectedUsername: string | undefined,
): AccountInfo | undefined {
  if (expectedUsername === undefined) {
    return accounts.length === 1 ? accounts[0] : undefined;
  }
  const normalizedExpected = expectedUsername.toLocaleLowerCase("en-US");
  return accounts.find(
    (account) =>
      account.username.toLocaleLowerCase("en-US") === normalizedExpected,
  );
}

export async function getAccessToken(
  mailbox: MailboxConfig,
  cacheDirectory: string,
  interactive: boolean,
): Promise<string> {
  const environment = readAuthEnvironment();
  const client = await createClient(mailbox, cacheDirectory, environment);
  const accounts = await client.getTokenCache().getAllAccounts();
  const account = findAccount(accounts, mailbox.expectedUsername);

  if (account !== undefined) {
    try {
      const result = await client.acquireTokenSilent({
        account,
        scopes: GRAPH_SCOPES,
      });
      return result.accessToken;
    } catch (error) {
      if (!interactive) {
        const message = error instanceof Error ? error.message : String(error);
        throw new Error(
          `Silent sign-in failed for ${mailbox.label}: ${message}. Run the auth command again.`,
        );
      }
    }
  } else if (!interactive) {
    throw new Error(
      `No cached sign-in exists for ${mailbox.label}. Run "npm run auth -- --mailbox ${mailbox.id}" first.`,
    );
  }

  const result = await client.acquireTokenByDeviceCode({
    scopes: GRAPH_SCOPES,
    deviceCodeCallback: (deviceCode) => {
      console.log(`\nSign in for ${mailbox.label}:`);
      console.log(deviceCode.message);
    },
  });
  if (result === null) {
    throw new Error(`Microsoft sign-in was cancelled for ${mailbox.label}.`);
  }

  const signedInUsername = result.account?.username;
  if (
    mailbox.expectedUsername !== undefined &&
    signedInUsername?.toLocaleLowerCase("en-US") !==
      mailbox.expectedUsername.toLocaleLowerCase("en-US")
  ) {
    throw new Error(
      `Signed in as ${signedInUsername ?? "an unknown account"}, but config expected ${mailbox.expectedUsername} for ${mailbox.label}.`,
    );
  }

  return result.accessToken;
}
