require("dotenv").config(); // Load env variables
const path = require("path");

const express = require("express");
const axios = require("axios");
const cors = require("cors");
const app = express();
const port = 3000;

// Middlewares
app.use(cors());
app.use(express.json());

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html")); // Serve index.html
});

// Azure/Power BI credentials
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const tenantId = process.env.TENANT_ID;
const workspaceId = process.env.WORKSPACE_ID;
const reportId = process.env.REPORT_ID;

// Debug log env vars (optional, safe to remove later)
console.log("CLIENT_ID:", clientId);
console.log("WORKSPACE_ID:", workspaceId);

// Get Azure access token
async function getAzureAccessToken() {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const data = new URLSearchParams();
  data.append("grant_type", "client_credentials");
  data.append("client_id", clientId);
  data.append("client_secret", clientSecret);
  data.append("scope", "https://analysis.windows.net/powerbi/api/.default");

  const response = await axios.post(url, data);
  return response.data.access_token;
}

// Get Power BI Embed Token
async function getEmbedToken(accessToken) {
  const url = `https://api.powerbi.com/v1.0/myorg/groups/${workspaceId}/reports/${reportId}/GenerateToken`;

  const requestBody = {
    accessLevel: "view",
  };

  const config = {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  };

  const response = await axios.post(url, requestBody, config);
  return response.data.token;
}

// Home route (for testing)
app.get("/", (req, res) => {
  res.send("Server running. Use GET /get-embed-token to generate token.");
});

// Endpoint to get Embed Token
app.all("/get-embed-token", async (req, res) => {
  try {
    const accessToken = await getAzureAccessToken();
    const embedToken = await getEmbedToken(accessToken);
    res.json({ embedToken });
  } catch (error) {
    console.error("Error:", error.response?.data || error.message);
    res.status(500).send("Error generating embed token.");
  }
});

// Start server
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}/get-embed-token`);
  // console.log(`Server running at http://localhost:${port}`);
});
