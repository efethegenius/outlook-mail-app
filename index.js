const express = require("express");
const session = require("express-session");
const cors = require("cors");
const app = express();
const { Client } = require("@microsoft/microsoft-graph-client");
const {
  PublicClientApplication,
  ConfidentialClientApplication,
} = require("@azure/msal-node");

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(
  session({
    secret: "any_secret_key",
    resave: false,
    saveUninitialized: false,
  })
);

let port = process.env.PORT || 3000;

const clientId = "your_client_id";
const clientSecret = "your_client_secret";
const tenantId = "your_tenant_id";
const redirectUri = "http://localhost:3000"; //or any redirect uri you set on the azure AD

const scopes = ["https://graph.microsoft.com/.default"];

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

app.get("/signin", (req, res) => {
  const authCodeUrlParameters = {
    scopes,
    redirectUri,
  };

  pca.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
    res.redirect(response);
  });
});

app.get("/", (req, res) => {
  const tokenRequest = {
    code: req.query.code,
    scopes,
    redirectUri,
    clientSecret: clientSecret,
  };

  pca
    .acquireTokenByCode(tokenRequest)
    .then((response) => {
      // Store the user-specific access token in the session for future use
      req.session.accessToken = response.accessToken;

      // Redirect the user to a profile page or any other secure route
      // This time, we are redirecting to the get-access-token route to generate a client token
      res.redirect("/get-access-token");
    })
    .catch((error) => {
      console.log(error);
      res.status(500).send(error);
    });
});

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

app.use("/get-mails/:num", async (req, res) => {
  const num = req.params.num;

  try {
    const userAccessToken = req.session.accessToken;
    const clientAccessToken = req.session.clientAccessToken;

    if (!userAccessToken) {
      return res
        .status(401)
        .send("User not authenticated. Please sign in first.");
    }

    if (!clientAccessToken) {
      return res
        .status(401)
        .send(
          "Client not authenticated. Please acquire the client access token first."
        );
    }

    const client = Client.init({
      authProvider: (done) => {
        done(null, userAccessToken);
      },
    });

    const messages = await client.api("/me/messages").top(num).get();
    res.send(messages);
  } catch (error) {
    res.status(500).send(error);
    console.log("Error fetching messages:", error.message);
  }
});

app.use("/send-mail/:recipient", async (req, res) => {
  const recipient = req.params.recipient;

  try {
    // Retrieve the user and client access tokens from the session
    const userAccessToken = req.session.accessToken;
    const clientAccessToken = req.session.clientAccessToken;

    // Check if the user and client are authenticated
    if (!userAccessToken) {
      return res
        .status(401)
        .send("User not authenticated. Please sign in first.");
    }

    if (!clientAccessToken) {
      return res
        .status(401)
        .send(
          "Client not authenticated. Please acquire the client access token first."
        );
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
