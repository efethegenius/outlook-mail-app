# outlook-mail-app
Connecting to Microsoft Outlook and fetching emails using Node.js using the Microsoft Graph API, which provides access to various Microsoft services, including Outlook, through a secure authentication process

**Introduction:**
In this tutorial, we will learn how to connect to Microsoft Outlook and fetch emails using Node.js. We will be using the Microsoft Graph API, which provides access to various Microsoft services, including Outlook, through a secure authentication process. By the end of this guide, you will have a functional Node.js application capable of fetching emails from your Outlook account.

**Prerequisites:**
Before we get started, make sure you have the following prerequisites in place:
1. Node.js installed on your machine.
2. A Microsoft Azure account and an active directory.

# Step 1: Setting Up the Project
Create a new directory for your Node.js project and navigate into it using the command line. Run the following commands to initialize the project and install required dependencies:

```

mkdir outlook-mail-app
cd outlook-mail-app
npm init -y
npm install express express-session cors @microsoft/microsoft-graph-client @azure/msal-node

```

**NB:** Create a new file in your app folder called index.js.

or you can set up your project by creating a new folder whatever directory you want, open the new folder vs code and running the following in the terminal:

```

npm init -y
npm install express express-session cors @microsoft/microsoft-graph-client @azure/msal-node

```
**NB:** Create a new file in your app folder called index.js.

# Step 2: Setting Up the Microsoft Azure App Registration
To interact with Microsoft Graph API, you need to register your application on the Azure portal with the following steps:

1. Go to the Azure portal (https://portal.azure.com) and sign in with your Microsoft account.
2. Navigate to the “Azure Active Directory” service and select “App registrations.”
3. Click on “New registration” and fill in the required details like “Name” and “Supported account types.”
4. In the “Redirect URI” section, add “http://localhost:3000" as the Redirect URI for this tutorial (you can modify it later for production use).
5. After registering, note down the “Application (client) ID” and “Directory (tenant) ID” in the overview page.

**Generating a Client Secret:**
1. Navigate to “Certificates & Secrets”

2. Click on “New client secret” and fill in the name and expiration duration of the secret token.

# Step 3: Configuring the Node.js Application

Navigate to your index.js file and follow through with the steps below.

Importing Required Modules:

```

const express = require("express");
const session = require("express-session");
const cors = require("cors");
const app = express();
const { Client } = require("@microsoft/microsoft-graph-client");
const { PublicClientApplication, ConfidentialClientApplication } = require("@azure/msal-node");

```

**Explanation:**

**express:** A popular Node.js framework for building web applications.
**session:** Middleware for managing and storing user sessions. We would use it to store the authentication tokens we generate from Microsoft Graph API.
**cors:** Middleware for enabling Cross-Origin Resource Sharing.
**Client:** A class from @microsoft/microsoft-graph-client used to interact with the Microsoft Graph API.
**PublicClientApplication and ConfidentialClientApplication:** Classes from @azure/msal-node for handling authentication.


**Configuring Express Middleware:**

```

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: "any_secret_key",
  resave: false,
  saveUninitialized: false,
}));

```
**Explanation:**

**cors():** Enables Cross-Origin Resource Sharing, allowing the application to be accessed from different domains.
**express.json():** Parses incoming requests with JSON payloads.
**express.urlencoded():** Parses incoming requests with URL-encoded payloads.
**session():** Configures the express session middleware with a secret for secure session management to store our authentication tokens.

**Application Constants and Configuration:**
```

let port = process.env.PORT || 3000;

const clientId = "your_client_id";
const clientSecret = "your_client_secret";
const tenantId = "your_tenant_id";
const redirectUri = "http://localhost:3000"; //or any redirect uri you set on the azure AD

const scopes = ["https://graph.microsoft.com/.default"];

```
**Explanation:**

**port:** Specifies the port on which the application will listen for incoming requests.
clientId, clientSecret, and tenantId: These are the obtained details from your Azure App Registration that will be used for authentication.

**redirectUri:** The URL to redirect users after authentication.
scopes: The permissions required for the application to access the Microsoft Graph API. The .defaultsignifies that the application will use all granted permissions in the App Registration.
Creating MSAL Clients:

```

const msalConfig = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri,
  },
};

const pca = new PublicClientApplication(msalConfig);

const ccaConfig = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientSecret,
  },
};

const cca = new ConfidentialClientApplication(ccaConfig);

```

**Explanation:**

**msalConfig:** Configuration object for the Public Client Application, this is used for user authentication.
pca: The Public Client Application instance created with the provided configuration.

**ccaConfig:** Configuration object for the Confidential Client Application, this is used for client authentication.
cca: The Confidential Client Application instance created with the provided configuration.

**Implementing User Authentication Flow:**
```

app.get("/signin", (req, res) => {
  const authCodeUrlParameters = {
    scopes,
    redirectUri,
  };

  pca.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
    res.redirect(response);
  });
});

```

**Explanation:**

When a user accesses /signin, the server generates the URL to redirect the user for authentication using pca.getAuthCodeUrl().
The user is redirected to the Microsoft login page to authenticate and provide consent for the requested permissions.

**Handling the Redirect After User Authentication:**
```

app.get("/", (req, res) => {
  const tokenRequest = {
    code: req.query.code,
    scopes,
    redirectUri,
    clientSecret: clientSecret,
  };

  pca.acquireTokenByCode(tokenRequest).then((response) => {
    // Store the user-specific access token in the session for future use
    req.session.accessToken = response.accessToken;

// Redirect the user to a profile page or any other secure route
// This time, we are redirecting to the get-access-token route to generate a client token
    res.redirect("/get-access-token"); 
  }).catch((error) => {
    console.log(error);
    res.status(500).send(error);
  });
});

```
**Explanation:**

After the user logs in and provides consent, they are redirected back to the root path /.
The server extracts the authorization code from the query parameters and uses it to request an access token for the user using pca.acquireTokenByCode().
The obtained access token is stored in the user’s session for future use.
The user is then automatically redirected to the get-access-token route. (next code) to generate the client access token.

**Obtaining Client Access Token:**
```

app.get("/get-access-token", async (req, res) => {
  try {
    const tokenRequest = {
      scopes,
      clientSecret: clientSecret,
    };

    const response = await cca.acquireTokenByClientCredential(tokenRequest);
    const accessToken = response.accessToken;

    // Store the client-specific access token in the session for future use
    req.session.clientAccessToken = accessToken; // This will now be stored in the session

    res.send("Access token acquired successfully!");
  } catch (error) {
    res.status(500).send(error);
    console.log("Error acquiring access token:", error.message);
  }
});

```

**Explanation:**

The route /get-access-token allows the application to obtain a client-specific access token (client access token) using cca.acquireTokenByClientCredential().
The client access token is essential for making calls to the Microsoft Graph API on behalf of the application. this might not be needed if you only want to make user specific calls.
Hurray! we are done with setting up the tokens needed to query our mailboxes. Next, we are going to write an endpoint to fetch mails from our mailbox.

**Fetching Emails:**
```

app.use("/get-mails/:num", async (req, res) => {
  const num = req.params.num;

  try {
    const userAccessToken = req.session.accessToken;
    const clientAccessToken = req.session.clientAccessToken;

    // Check if the user and client are authenticated
    if (!userAccessToken) {
      return res.status(401).send("User not authenticated. Please sign in first.");
    }

    if (!clientAccessToken) {
      return res.status(401).send("Client not authenticated. Please acquire the client access token first.");
    }

    // Initialize the Microsoft Graph API client using the user access token
    const client = Client.init({
      authProvider: (done) => {
        done(null, userAccessToken);
      },
    });

    // Fetch the user's emails using the Microsoft Graph API
    const messages = await client.api("/me/messages").top(num).get();
    res.send(messages);
  } catch (error) {
    res.status(500).send(error);
    console.log("Error fetching messages:", error.message);
  }
});

```

**Explanation:**

The route /get-mails/:num fetches emails for the authenticated user using the Microsoft Graph API.
The num parameter specifies the number of emails to fetch.
The function checks if both the user and client are authenticated by verifying the presence of their respective access tokens in the session.
It uses the Client.init() method to create a Microsoft Graph API client with the user access token.
The client then makes an API call to fetch the user’s emails using client.api("/me/messages").top(num).get().

# Step 4: Implementing the Authentication Flow
The provided code demonstrates the authentication flow using the Microsoft Authentication Library (MSAL) for Node.js. It uses both the Public Client Application (for user authentication) and Confidential Client Application (for client authentication). Make sure to replace the placeholders for `clientId`, `clientSecret`, and `tenantId` with the ones you obtained from your Azure App Registration.

# Step 5: Obtaining Required Permissions
For this tutorial, we will fetch emails from the user’s mailbox. To do this, we need the `Mail.Read` application permission, which requires admin consent in a production environment. For testing purposes, you can grant admin consent for yourself through the Azure portal:
1. In the App Registration page, go to “API permissions.”
2. Click on “Add a permission” and choose “Microsoft Graph.”
3. Select “Application permissions.”
4. Select “Microsoft Graph” -> “Delegated Permissions”. In the “Select permissions” search box, type “Mail.Read,” and then check the box next to “Mail.Read.”
5. Click on “Add permissions” and then “Grant admin consent.”

# Step 6: Running the Application
Now that we have configured everything, run the Node.js application using the following command:
```

node index.js

```

# Step 7: Testing the Application

Open your web browser and navigate to http://localhost:3000/signin . This will redirect you to the Microsoft login page. Sign in with your Outlook account to authenticate the application. If you are authenticated successfully, you should see a message on the screen that reads Access token acquired successfully!.

# BONUS CODE
**Sending an Email with our tokens:**

```

app.use("/send-mail/:recipient", async (req, res) => {
  const recipient = req.params.recipient;

  try {
    // Retrieve the user and client access tokens from the session
    const userAccessToken = req.session.accessToken;
    const clientAccessToken = req.session.clientAccessToken;

    // Check if the user and client are authenticated
    if (!userAccessToken) {
      return res.status(401).send("User not authenticated. Please sign in first.");
    }

    if (!clientAccessToken) {
      return res.status(401).send("Client not authenticated. Please acquire the client access token first.");
    }

    // Initialize the Microsoft Graph API client using the user access token
    const client = Client.init({
      authProvider: (done) => {
        done(null, userAccessToken);
      },
    });

    // Prepare the email data
    const sendMail = {
      message: {
        subject: "Wanna go out for lunch?",
        body: {
          contentType: "Text",
          content: "I know a sweet spot that just opened around us!",
        },
        toRecipients: [
          {
            emailAddress: {
              address: recipient,
            },
          },
        ],
      },
      saveToSentItems: false,
    };

    // Send the email using the Microsoft Graph API
    const response = await client.api("/me/sendMail").post(sendMail);
    res.send(response);
  } catch (error) {
    res.status(500).send(error);
    console.log("Error sending message:", error.message);
  }
});


app.listen(port, () => {
  console.log(`app listening on port ${port}`);
});

```

**Explanation:**

The route /send-mail/:recipient allows the application to send an email to the specified recipient using the Microsoft Graph API.
The recipient parameter specifies the email address of the recipient.
Similar to fetching emails, the function checks if both the user and client are authenticated before proceeding.
The function uses the Client.init() method to create a Microsoft Graph API client with the user access token.
It prepares the email data in the sendMail object and sends the email using client.api("/me/sendMail").post(sendMail).
To do this, we need the `Mail.Send` application permission, which requires admin consent in a production environment. We can grant the permission just like we did for `Mail.Read` above.
To send a mail, navigate to your browser and input http:localhost:3000/send-mail/your-recipient-email@address.com . this would send a mail to the email address specified on the URL.

NB: The body and header texts are hard coded, you can expand the project to have inputs to type in the header and body.

**Conclusion:**
Congratulations! You have successfully connected to Microsoft Outlook and fetched emails using Node.js. You have learned how to set up a Node.js project, authenticate with Microsoft Graph API, and use the acquired access tokens to fetch emails from your Outlook account. By following this guide, you can build a functional Node.js application to manage Outlook emails programmatically.

Remember that this is just a starting point, and you can expand this application to perform various other tasks with Outlook emails using the Microsoft Graph API. Happy coding!
