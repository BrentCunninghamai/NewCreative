# Code signing (remove the SmartScreen warning)

The published `M365Migrate.exe` is **unsigned** by default, so Windows SmartScreen
shows "Windows protected your PC" on first run (click *More info → Run anyway*).
Signing the binary removes that prompt and builds reputation over time.

The CI pipeline is **already wired to sign** — it just needs credentials. Until
they're configured the signing step is skipped automatically and builds stay
green. Two things are required, and only you can do them (signing requires
identity/organization validation tied to a legal entity, and payment):

## Recommended: Azure Trusted Signing (~$10/month)

This is the modern, low-cost path and is what the pipeline uses
(`azure/trusted-signing-action`). High level:

1. In the Azure portal, create a **Trusted Signing account** (Microsoft.CodeSigning).
2. Complete the **identity validation** for your organization (Microsoft verifies
   you; this is the part only you can do, and it can take a few days).
3. Create a **Certificate profile** (Public Trust) under the account.
4. Create an **App registration** (service principal) and grant it the
   **Trusted Signing Certificate Profile Signer** role on the account.
5. Add these **GitHub Actions secrets** to the repository
   (Settings → Secrets and variables → Actions):

   | Secret | Value |
   |--------|-------|
   | `AZURE_TENANT_ID` | the app registration's directory (tenant) id |
   | `AZURE_CLIENT_ID` | the app registration's application (client) id |
   | `AZURE_CLIENT_SECRET` | a client secret for that app registration |
   | `SIGNING_ENDPOINT` | your region endpoint, e.g. `https://eus.codesigning.azure.net/` |
   | `SIGNING_ACCOUNT` | the Trusted Signing account name |
   | `SIGNING_PROFILE` | the certificate profile name |

Once `AZURE_CLIENT_ID` and `SIGNING_ACCOUNT` are present, the next push to `main`
signs the exe automatically and the rolling release ships a **signed** binary.

## Alternative: a traditional OV/EV code-signing certificate

You can instead buy an OV or EV code-signing certificate from a CA (Sectigo,
DigiCert, etc.). Modern OV/EV certs are issued on a hardware token or cloud HSM,
which complicates unattended CI signing — that's why Azure Trusted Signing is
recommended here. If you prefer this route, tell me and I'll adapt the pipeline
to your CA's signing method.

## I can't do this part for you

Buying a certificate / setting up Trusted Signing requires payment and identity
verification of your organization. Once you've added the secrets above (or have a
cert), I'll confirm the pipeline signs correctly and the SmartScreen prompt is
gone.
