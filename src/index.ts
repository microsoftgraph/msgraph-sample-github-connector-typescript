// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import 'dotenv-flow/config';
import * as readline from 'readline-sync';
import { ExternalConnectors } from '@microsoft/microsoft-graph-types-beta';
import { marked } from 'marked';

import { MenuChoice, menuPrompts, ItemTypeChoice, itemTypes } from './menu';
import PlainTextRenderer from './markdown/plainTextRenderer';
import SearchConnectorService from './services/searchConnectorService';
import RepositoryService, {
  Issue,
  IssueComment,
  IssueEvent,
  RepoEvent,
  Repository,
} from './services/repositoryService';
import M365AppConfigService from './services/m365AppConfigService';

async function main() {
  // Connector service for making Microsoft Graph
  // calls to manage connector
  const connectorService = new SearchConnectorService({
    tenantId: process.env.TENANT_ID,
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    gitHubOwner: process.env.GITHUB_REPO_OWNER,
    gitHubRepo: process.env.GITHUB_REPO,
    placeHolderUserId: process.env.PLACEHOLDER_USER_ID,
  });

  // Repo service for getting information from GitHub
  const repoService = new RepositoryService({
    gitHubOwner: process.env.GITHUB_REPO_OWNER,
    gitHubRepo: process.env.GITHUB_REPO,
    gitHubToken: process.env.GITHUB_TOKEN,
  });

  // Check for simplified admin switch
  if (process.argv.includes('--use-simplified-admin')) {
    // Start listener service
    const m365ConfigService = new M365AppConfigService(connectorService, {
      clientId: process.env.CLIENT_ID,
      tenantId: process.env.TENANT_ID,
      port: parseInt(process.env.PORT_NUMBER ?? '5001'),
    });

    m365ConfigService.listen();
  } else {
    // Run interactively
    await runInteractivelyAsync(connectorService, repoService);
  }
}

/**
 * Present menu to user and process their choice.
 *
 * @param connectorService - The connector service.
 * @param repoService - The repository service.
 */
async function runInteractivelyAsync(
  connectorService: SearchConnectorService,
  repoService: RepositoryService,
) {
  let choice: MenuChoice = MenuChoice.Invalid;
  let currentConnection: ExternalConnectors.ExternalConnection | undefined =
    undefined;

  while (choice !== MenuChoice.Exit) {
    console.log(`Current connection: ${currentConnection?.name ?? 'NONE'}`);
    choice = readline.keyInSelect(menuPrompts, 'Select an option', {
      cancel: 'Exit',
    });

    switch (choice) {
      case MenuChoice.Exit:
        console.log('Goodbye...');
        break;
      case MenuChoice.CreateConnection:
        currentConnection =
          await createConnectionInteractivelyAsync(connectorService);
        break;
      case MenuChoice.SelectConnection:
        currentConnection =
          await selectConnectionInteractivelyAsync(connectorService);
        break;
      case MenuChoice.DeleteConnection:
        if (currentConnection) {
          await deleteConnectionInteractivelyAsync(
            connectorService,
            currentConnection.id,
          );
          currentConnection = undefined;
        } else {
          console.log(
            'No connection selected. Please create a new connection or select an existing connection.',
          );
        }
        break;
      case MenuChoice.RegisterSchema:
        if (currentConnection) {
          await registerSchemaInteractivelyAsync(
            connectorService,
            currentConnection.id,
          );
        } else {
          console.log(
            'No connection selected. Please create a new connection or select an existing connection.',
          );
        }
        break;
      case MenuChoice.PushAllItems:
        if (currentConnection) {
          await pushItemsInteractivelyAsync(
            connectorService,
            repoService,
            currentConnection.id,
          );
        } else {
          console.log(
            'No connection selected. Please create a new connection or select an existing connection.',
          );
        }
        break;
      default:
        console.log('Invalid choice!');
    }
  }
}

/**
 * Prompt the user for information to create a new connection.
 *
 * @param connectorService - The connector service.
 * @returns The created connection.
 */
async function createConnectionInteractivelyAsync(
  connectorService: SearchConnectorService,
): Promise<ExternalConnectors.ExternalConnection | undefined> {
  // Prompt for connection ID
  const connectionId = readline.question(
    'Enter a unique ID for the new connection (3-32 alphanumeric characters): ',
    {
      limit: /^[0-9a-zA-Z]{3,32}$/,
      limitMessage: 'ID must be alphanumeric and 3 to 32 characters long',
    },
  );

  // Prompt for connection name
  const connectionName = readline.question(
    'Enter a name for the new connection: ',
    {
      limit: function (path) {
        return path.length > 0;
      },
      limitMessage: 'Name is required',
    },
  );

  // Prompt for a description
  const connectionDescription = readline.question(
    'Enter a description for the new connection (OPTIONAL): ',
  );

  // Prompt for type of data connection will support
  const itemType: ItemTypeChoice = readline.keyInSelect(
    itemTypes,
    'What type of data?',
  );

  try {
    const connection = await connectorService.createConnectionAsync(
      connectionId,
      connectionName,
      itemType,
      connectionDescription,
    );
    console.log(
      `New connection created - Name: ${connection?.name}, Id: ${connection?.id}`,
    );
    return connection;
  } catch (error) {
    console.log(`Error creating connection: ${JSON.stringify(error, null, 2)}`);
  }
}

/**
 * Get existing connections and prompt the user to choose one.
 *
 * @param connectorService - The connector service.
 * @returns The selected connection.
 */
async function selectConnectionInteractivelyAsync(
  connectorService: SearchConnectorService,
): Promise<ExternalConnectors.ExternalConnection | undefined> {
  console.log('Getting existing connections...');

  try {
    const response = await connectorService.getConnectionsAsync();
    const connections =
      response.value as ExternalConnectors.ExternalConnection[];
    if (connections.length <= 0) {
      console.log('No connections exist. Please create a new connection.');
      return;
    }

    const connectionNames = connections.map(
      (c) => c.name?.toString() ?? 'No name',
    );
    const selectedIndex = readline.keyInSelect(
      connectionNames,
      'Choose one of the following connections',
    );

    if (selectedIndex >= 0) {
      return connections[selectedIndex];
    }
  } catch (error) {
    console.log(`Error getting connections: ${JSON.stringify(error, null, 2)}`);
  }
}

/**
 * Prompt user to confirm, then delete current connection.
 *
 * @param connectorService - The connector service.
 * @param connectionId - The ID of the current connection.
 */
async function deleteConnectionInteractivelyAsync(
  connectorService: SearchConnectorService,
  connectionId?: string,
) {
  if (readline.keyInYNStrict()) {
    try {
      await connectorService.deleteConnectionAsync(connectionId);
      console.log('Connection deleted successfully.');
    } catch (error) {
      console.log(
        `Error deleting connection: ${JSON.stringify(error, null, 2)}`,
      );
    }
  }
}

/**
 * Prompt the user for the type of data then register the appropriate schema.
 *
 * @param connectorService - The connector service.
 * @param connectionId - The ID of the current connection.
 */
async function registerSchemaInteractivelyAsync(
  connectorService: SearchConnectorService,
  connectionId?: string,
) {
  if (!connectionId) {
    throw new Error('connectionId cannot be empty or undefined');
  }

  const itemType: ItemTypeChoice = readline.keyInSelect(
    itemTypes,
    'What type of data?',
  );

  console.log('Registering schema, this may take some time...');
  try {
    await connectorService.registerSchemaAsync(
      connectionId,
      itemType == ItemTypeChoice.Issues
        ? SearchConnectorService.issuesSchema
        : SearchConnectorService.reposSchema,
    );
    console.log('Schema registered successfully.');
  } catch (error) {
    console.log(`Error registering schema: ${JSON.stringify(error, null, 2)}`);
  }
}

/**
 * Prompt the user for the type of data then push data from GitHub to the connection.
 *
 * @param connectorService - The connector service.
 * @param repoService - The repository service.
 * @param connectionId - The ID of the current connection.
 */
async function pushItemsInteractivelyAsync(
  connectorService: SearchConnectorService,
  repoService: RepositoryService,
  connectionId?: string,
) {
  if (!connectionId) {
    throw new Error('connectionId cannot be empty or undefined');
  }

  const itemType: ItemTypeChoice = readline.keyInSelect(
    itemTypes,
    'What type of data?',
  );

  if (itemType === ItemTypeChoice.Issues) {
    await pushAllIssuesWithActivitiesAsync(
      connectorService,
      repoService,
      connectionId,
    );
  } else {
    await pushAllRepositoriesAsync(connectorService, repoService, connectionId);
  }
}

/**
 * Get open issues from configured GitHub repo and push to current connection.
 *
 * @param connectorService - The connector service.
 * @param repoService - The repository service.
 * @param connectionId - The ID of the current connection.
 */
async function pushAllIssuesWithActivitiesAsync(
  connectorService: SearchConnectorService,
  repoService: RepositoryService,
  connectionId: string,
) {
  let issues: Issue[] | undefined = undefined;
  try {
    issues = await repoService.getIssuesForRepositoryAsync();
  } catch (error) {
    console.log(`Error getting issues: ${JSON.stringify(error, null, 2)}`);
  }

  // Markdown to plain text renderer
  const plainText = new PlainTextRenderer();

  if (issues) {
    for (const issue of issues) {
      console.log(`Adding/updating issue ${issue.number}`);

      let issueEvents: IssueEvent[] = [];
      try {
        issueEvents = await repoService.getEventsForIssueAsync(issue.number);
      } catch (error) {
        console.log(
          `Error getting events for issue: ${JSON.stringify(error, null, 2)}`,
        );
      }

      let comments: IssueComment[] = [];
      try {
        comments = await repoService.getCommentsForIssueAsync(issue.number);
      } catch (error) {
        console.log(
          `Error getting comments for issue: ${JSON.stringify(error, null, 2)}`,
        );
      }

      try {
        const issueItem =
          await connectorService.createExternalItemFromIssueAsync(
            issue,
            issueEvents,
          );

        // Generate content for the issue by concatenating
        // the body of the issue + all comments
        let issueContent = marked.parse(issue.body || '', {
          renderer: plainText,
        });

        for (const comment of comments) {
          issueContent += `\n${marked.parse(comment.body || '', {
            renderer: plainText,
          })}`;
        }

        issueItem.content = {
          type: 'text',
          value: issueContent,
        };

        await connectorService.addOrUpdateItemAsync(connectionId, issueItem);

        const activities =
          await connectorService.createExternalActivitiesFromIssueEventsAsync(
            issueEvents,
          );
        await connectorService.addIssueActivitiesAsync(
          connectionId,
          issue.number.toString(),
          activities,
        );
        console.log('DONE');
      } catch (error) {
        console.log(
          `Error adding/updating issue: ${JSON.stringify(error, null, 2)}`,
        );
      }
    }
  }
}

/**
 * Get repositories from configured GitHub owner and push to current connection.
 *
 * @param connectorService - The connector service.
 * @param repoService - The repository service.
 * @param connectionId - The ID of the current connection.
 */
async function pushAllRepositoriesAsync(
  connectorService: SearchConnectorService,
  repoService: RepositoryService,
  connectionId: string,
) {
  let repos: Repository[] | undefined = undefined;
  try {
    repos = await repoService.getRepositoriesAsync();
  } catch (error) {
    console.log(
      `Error getting repositories: ${JSON.stringify(error, null, 2)}`,
    );
  }

  // Markdown to plain text renderer
  const plainText = new PlainTextRenderer();

  if (repos) {
    for (const repo of repos) {
      console.log(`Adding/updating repository ${repo.name}...`);

      let repoEvents: RepoEvent[] = [];
      try {
        repoEvents = await repoService.getEventsForRepoAsync(repo.name);
      } catch (error) {
        console.log(
          `Error getting events for repo: ${JSON.stringify(error, null, 2)}`,
        );
      }

      const repoItem = await connectorService.createExternalItemFromRepoAsync(
        repo,
        repoEvents,
      );

      if (repo.visibility === 'public') {
        // For public repositories,
        // set content to the README
        const readme = await repoService.getReadmeAsync(repo.name);

        if (readme) {
          const readmeContent = Buffer.from(readme.content, 'base64').toString(
            'utf8',
          );
          const plainContent = marked.parse(readmeContent, {
            renderer: plainText,
          });

          repoItem.content = {
            type: 'text',
            value: plainContent,
          };
        }
      } else {
        // For private repositories,
        // set content to the JSON representation
        repoItem.content = {
          type: 'text',
          value: JSON.stringify(repo),
        };
      }

      try {
        await connectorService.addOrUpdateItemAsync(connectionId, repoItem);
        console.log('DONE');
      } catch (error) {
        console.log(
          `Error adding/updating repository: ${JSON.stringify(error, null, 2)}`,
        );
      }
    }
  }
}

main();
