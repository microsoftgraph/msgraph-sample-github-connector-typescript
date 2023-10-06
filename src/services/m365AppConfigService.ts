// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as express from 'express';
import Router from 'express-promise-router';
import { JwtHeader, SigningKeyCallback, verify } from 'jsonwebtoken';
import { JwksClient } from 'jwks-rsa';
import {
  ChangeNotificationCollection,
  ExternalConnectors,
} from '@microsoft/microsoft-graph-types-beta';

import SearchConnectorService from './searchConnectorService';
import ConnectorData from '../types/connectorData';
import { ItemTypeChoice } from '../menu';

export type M365AppConfigServiceOptions = {
  /**
   * The "Application (client) ID" of the app registration in Azure.
   */
  clientId?: string;
  /**
   * The "Directory (tenant) ID" of the app registration in Azure.
   */
  tenantId?: string;
  /**
   * The port number to listen on.
   */
  port: number;
};

export default class M365AppConfigService {
  private app: express.Express;
  private clientId: string;
  private tenantId: string;
  private port: number;

  private connectorService: SearchConnectorService;

  constructor(
    connectorService: SearchConnectorService,
    options: M365AppConfigServiceOptions,
  ) {
    if (!options.tenantId || !options.clientId) {
      throw new Error('Invalid app registration details, please see README');
    }

    this.clientId = options.clientId;
    this.tenantId = options.tenantId;
    this.port = options.port;
    this.connectorService = connectorService;

    this.app = express();
    this.app.use(express.json());
    const router = Router();
    this.app.use(router);
    this.app.locals.configService = this;

    router.post('/', this.processRequest);
  }

  /**
   * Start the Express app to listen on the specified port.
   */
  public listen() {
    this.app.listen(this.port, 'localhost', () => {
      console.log(`Server running at http://localhost:${this.port}`);
    });
  }

  /**
   * Function that processes POST requests.
   *
   * @param req - The incoming request.
   * @param res - The outgoing response.
   */
  private async processRequest(req: express.Request, res: express.Response) {
    const changeNotifications = req.body as ChangeNotificationCollection;
    const configService = res.app.locals.configService as M365AppConfigService;

    // Return 202 so Microsoft Graph won't retry notification
    res.sendStatus(202);

    if (
      changeNotifications &&
      changeNotifications.value &&
      changeNotifications.validationTokens
    ) {
      // Validate the validation tokens
      const validationResults = await Promise.all(
        changeNotifications.validationTokens.map((token) =>
          configService.isTokenValid(token),
        ),
      );

      const areTokensValid = validationResults.reduce((x, y) => x && y);
      if (areTokensValid) {
        for (const notification of changeNotifications.value) {
          // Process the resourceData object
          const connectorData = notification.resourceData as ConnectorData;
          if (
            connectorData &&
            connectorData['@odata.type'].toLowerCase() ===
              '#microsoft.graph.connector'
          ) {
            console.log(
              `Checking for existence of connection with connector ID: ${connectorData.id}`,
            );

            // Get all existing connections
            const allConnections =
              await configService.connectorService.getConnectionsAsync();
            let existingConnection:
              | ExternalConnectors.ExternalConnection
              | undefined = undefined;
            // Search for a connection with a matching connector ID
            for (const connection of allConnections.value as ExternalConnectors.ExternalConnection[]) {
              if (connection.connectorId === connectorData.id) {
                existingConnection = connection;
                break;
              }
            }

            if (connectorData.state === 'enabled') {
              // Request is to enable the connector. If it already exists,
              // do nothing. Otherwise create a new connection.
              console.log('Received request to enable connector');
              if (existingConnection) {
                console.log('Connection already exists');
              } else {
                try {
                  await configService.connectorService.createConnectionAsync(
                    'GitHubIssuesM365',
                    'GitHub Issues for M365 App',
                    ItemTypeChoice.Issues,
                    'This connector was created by an M365 app',
                    connectorData.connectorsTicket,
                    connectorData.id,
                  );
                  console.log('Created connection successfully');
                  console.log('Registering schema, this may take some time...');

                  await configService.connectorService.registerSchemaAsync(
                    'GitHubIssuesM365',
                    SearchConnectorService.issuesSchema,
                  );
                  console.log('Registered schema successfully');
                } catch (error) {
                  console.log(
                    `Error creating connection: ${JSON.stringify(
                      error,
                      null,
                      2,
                    )}`,
                  );
                }
              }
            } else {
              // Request is to disable the connector. If a connection
              // exists, delete it. Otherwise do nothing.
              console.log('Received request to disable connector');
              if (existingConnection) {
                try {
                  await configService.connectorService.deleteConnectionAsync(
                    existingConnection.id,
                  );
                  console.log('Connection deleted successfully');
                } catch (error) {
                  console.log(
                    `Error deleting connection: ${JSON.stringify(
                      error,
                      null,
                      2,
                    )}`,
                  );
                }
              }
            }
          }
        }
      }
    }
  }

  /**
   * Validates a signed JWT.
   *
   * @param token - The token to validate.
   * @returns A value indicating if the token is valid.
   */
  public async isTokenValid(token: string): Promise<boolean> {
    return new Promise((resolve) => {
      const options = {
        audience: [this.clientId],
        issuer: [
          `https://login.microsoftonline.com/${this.tenantId}/v2.0`,
          `https://sts.windows.net/${this.tenantId}/`,
        ],
      };

      verify(token, this.getKey, options, (error) => {
        if (error) {
          console.log(`Token validation error: ${error.message}`);
          resolve(false);
        } else {
          resolve(true);
        }
      });
    });
  }

  /**
   * Gets JWT signing keys from Microsoft identity.
   *
   * @param header - The JWT header from the token to validate.
   * @param callback - The callback method to call once the key has been retrieved.
   */
  private async getKey(header: JwtHeader, callback: SigningKeyCallback) {
    // Configure JSON web key set client to get keys
    // from well-known Microsoft identity endpoint
    const jwksClient = new JwksClient({
      jwksUri: 'https://login.microsoftonline.com/common/discovery/v2.0/keys',
    });

    try {
      const key = await jwksClient.getSigningKey(header.kid);
      const signingKey = key.getPublicKey();
      callback(null, signingKey);
    } catch (error) {
      callback(new Error(JSON.stringify(error)));
    }
  }
}
