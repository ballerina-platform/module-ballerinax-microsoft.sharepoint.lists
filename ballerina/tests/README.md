# Running Tests

## Prerequisites

Before running the tests, ensure you have configured the necessary credentials and environment for the `microsoft.sharepoint.lists` connector. Refer to the connector setup guide for detailed instructions:

[microsoft.sharepoint.lists Connector README](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/blob/main/ballerina/README.md)

## Running Tests

The tests for this connector require valid Microsoft SharePoint credentials and run directly against the live SharePoint service.

### Configure Test Credentials

Create a `Config.toml` file in the `tests` directory with the following content:

```toml
accessKeyId = "<your-access-key-id>"
secretAccessKey = "<your-secret-access-key>"
region = "<your-region>"
```

> **Note:** Replace each placeholder value with your actual Microsoft SharePoint credentials.

Alternatively, you can provide the credentials using environment variables.

**Linux/macOS:**

```bash
export accessKeyId="<your-access-key-id>"
export secretAccessKey="<your-secret-access-key>"
export region="<your-region>"
```

**Windows (Command Prompt):**

```cmd
set accessKeyId=<your-access-key-id>
set secretAccessKey=<your-secret-access-key>
set region=<your-region>
```

**Windows (PowerShell):**

```powershell
$env:accessKeyId = "<your-access-key-id>"
$env:secretAccessKey = "<your-secret-access-key>"
$env:region = "<your-region>"
```

### Execute the Tests

Once credentials are configured, run the tests from the `tests` directory using the following command:

```bash
bal test
```

### Skipping Live Tests

If you need to skip the live tests (for example, in environments where SharePoint access is unavailable), set the `LIVE_TEST_DISABLED` environment variable:

**Linux/macOS:**

```bash
export LIVE_TEST_DISABLED=true
bal test
```

**Windows (Command Prompt):**

```cmd
set LIVE_TEST_DISABLED=true
bal test
```

**Windows (PowerShell):**

```powershell
$env:LIVE_TEST_DISABLED = "true"
bal test
```
