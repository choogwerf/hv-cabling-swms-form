const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');

/**
 * Azure Function: Upload a PDF to SharePoint using Microsoft Graph.
 *
 * Expected request body (application/json):
 * {
 *   "filename": "Example.pdf",
 *   "fileContent": "<base64-encoded PDF>"
 * }
 *
 * Environment variables required:
 *   TENANT_ID   – Azure AD tenant ID
 *   CLIENT_ID   – Azure AD application (client) ID with Graph permissions
 *   CLIENT_SECRET – Client secret for the above application
 *   SITE_ID     – SharePoint site ID (from get_site API)
 *   DRIVE_ID    – Drive ID of the document library (can be default drive)
 *   FOLDER_PATH – Path within the drive to save files, e.g. 'SWMS Submissions'
 */

function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.TENANT_ID,
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
  );

  const authProvider = {
    getAccessToken: async () => {
      const token = await credential.getToken('https://graph.microsoft.com/.default');
      return token.token;
    }
  };

  return Client.initWithMiddleware({ authProvider });
}

module.exports = async function (context, req) {
  try {
    if (!req.body || !req.body.filename || !req.body.fileContent) {
      context.res = {
        status: 400,
        body: 'Request must include filename and fileContent.'
      };
      return;
    }

    const { filename, fileContent } = req.body;
    const buffer = Buffer.from(fileContent, 'base64');

    const client = getGraphClient();

    // Compose the upload path: /sites/{site-id}/drive/items/{parent-id}:/{filename}:/content
    // Use DRIVE_ID for the drive and FOLDER_PATH for folder path. If FOLDER_PATH is empty,
    // the file is saved at the root of the drive.
    const siteId = process.env.SITE_ID;
    const driveId = process.env.DRIVE_ID;
    const folderPath = process.env.FOLDER_PATH ? `${process.env.FOLDER_PATH}/` : '';

    const uploadUrl = `/drives/${driveId}/root:/${folderPath}${filename}:/content`;

    await client
      .api(uploadUrl)
      .put(buffer);

    context.res = {
      status: 200,
      body: 'File uploaded successfully.'
    };
  } catch (error) {
    context.res = {
      status: 500,
      body: `Upload failed: ${error.message}`
    };
  }
};
