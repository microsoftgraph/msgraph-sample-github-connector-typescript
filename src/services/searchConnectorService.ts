// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import 'isomorphic-fetch';
import { Response } from 'node-fetch';
import { ClientSecretCredential } from '@azure/identity';
import {
  Client,
  PageCollection,
  ResponseType,
} from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { ExternalConnectors } from '@microsoft/microsoft-graph-types-beta';
import { issuesSchema, reposSchema } from './schemas';
import { ItemTypeChoice } from '../menu';
import ItemIdResolverWithType, {
  itemIdResolverType,
} from '../types/itemIdResolverWithType';
import {
  Assignee,
  Issue,
  IssueEvent,
  Labels,
  RepoEvent,
  Repository,
} from './repositoryService';
import ExternalActivityWithType, {
  externalActivityType,
} from '../types/externalActivityWithType';
import { readFileSync } from 'fs';

export type SearchConnectorServiceOptions = {
  /**
   * The "Application (client) ID" of the app registration in Azure.
   */
  clientId?: string;
  /**
   * The "Directory (tenant) ID" of the app registration in Azure.
   */
  tenantId?: string;
  /**
   * The client secret of the app registration in Azure.
   */
  clientSecret?: string;
  /**
   * The GitHub user or organization.
   */
  gitHubOwner?: string;
  /**
   * The GitHub repository to ingest issues from.
   */
  gitHubRepo?: string;
  /**
   * The placeholder user ID to map to GitHub user logins.
   */
  placeHolderUserId?: string;
};

export default class SearchConnectorService {
  /**
   * Schema for ingesting GitHub issues.
   */
  static readonly issuesSchema: ExternalConnectors.Schema = issuesSchema;
  /**
   * Schema for ingesting GitHub repositories.
   */
  static readonly reposSchema: ExternalConnectors.Schema = reposSchema;

  private graphClient: Client;
  private gitHubOwner: string;
  private gitHubRepo: string;
  private placeHolderUserId: string;

  /**
   * Initializes a new instance of the SearchConnectorServiceOptions class.
   *
   * @param options - Contains the options for the class.
   */
  constructor(options: SearchConnectorServiceOptions) {
    if (!options.tenantId || !options.clientId || !options.clientSecret) {
      throw new Error('Invalid app registration details, please see README');
    }

    if (!options.gitHubOwner || !options.gitHubRepo) {
      throw new Error('Invalid GitHub details, please see README');
    }

    if (!options.placeHolderUserId) {
      throw new Error('Missing placeholder user ID, please see README');
    }

    this.gitHubOwner = options.gitHubOwner;
    this.gitHubRepo = options.gitHubRepo;
    this.placeHolderUserId = options.placeHolderUserId;

    const credential = new ClientSecretCredential(
      options.tenantId,
      options.clientId,
      options.clientSecret,
    );
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default'],
    });

    this.graphClient = Client.initWithMiddleware({
      authProvider: authProvider,
      defaultVersion: 'beta',
    });
  }

  /**
   * Creates a new connection.
   *
   * @param connectionId - The connection ID for the new connection.
   * @param name - The display name of the new connection.
   * @param itemType - The item type for the new connection (`issues` or `repos`).
   * @param description - The description of the new connection.
   * @param connectorTicket - The connector ticket when creating a connection from an M365 app.
   * @param connectorId - The connector ID when creating a connection from an M365 app.
   * @returns The new connection.
   */
  public async createConnectionAsync(
    connectionId: string,
    name: string,
    itemType: ItemTypeChoice,
    description?: string,
    connectorTicket?: string,
    connectorId?: string,
  ): Promise<ExternalConnectors.ExternalConnection | undefined> {
    const itemIdResolver: ItemIdResolverWithType = {
      '@odata.type': itemIdResolverType,
      priority: 1,
      itemId: itemType == ItemTypeChoice.Issues ? '{issueId}' : '{repo}',
      urlMatchInfo: {
        urlPattern:
          itemType == ItemTypeChoice.Issues
            ? `/${this.gitHubOwner}/${this.gitHubRepo}/issues/(?<issueId>[0-9]+)`
            : `/${this.gitHubOwner}/(?<repo>.*)/`,
        baseUrls: ['https://github.com'],
      },
    };

    const resultTemplateLayout = this.getResultTemplate(
      itemType == ItemTypeChoice.Issues
        ? './result-cards/result-typeIssues.json'
        : './result-cards/result-typeRepos.json',
    );

    const newConnection: ExternalConnectors.ExternalConnection = {
      id: connectionId,
      name: name,
      description: description,
      activitySettings: {
        urlToItemResolvers: [itemIdResolver],
      },
      searchSettings: {
        searchResultTemplates: [
          {
            id:
              itemType == ItemTypeChoice.Issues
                ? 'issueDisplay'
                : 'repoDisplay',
            priority: 1,
            layout: resultTemplateLayout,
          },
        ],
      },
    };

    const createRequest = this.graphClient.api('/external/connections');

    // Only send the M365 app properties (ConnectorId and GraphConnectors-Ticket header)
    // if they are both provided, otherwise API call will fail
    // See https://learn.microsoft.com/graph/connecting-external-content-deploy-teams
    if (connectorTicket && connectorId) {
      newConnection.connectorId = connectorId;
      createRequest.header('GraphConnectors-Ticket', connectorTicket);
    }

    return createRequest.post(newConnection);
  }

  /**
   * Gets existing connections.
   * @returns a PageCollection containing existing connections.
   */
  public async getConnectionsAsync(): Promise<PageCollection> {
    return this.graphClient.api('/external/connections').get();
  }

  /**
   * Deletes a connection.
   *
   * @param connectionId - The connection ID of the connection to delete.
   * @returns A Promise indicating the status of the asynchronous delete operation.
   */
  public async deleteConnectionAsync(connectionId?: string): Promise<void> {
    if (connectionId) {
      return this.graphClient
        .api(`/external/connections/${connectionId}`)
        .delete();
    }
  }

  /**
   * Registers a schema for a connection.
   *
   * @param connectionId - The connection ID of the connection.
   * @param schema - The schema to register.
   */
  public async registerSchemaAsync(
    connectionId: string,
    schema: ExternalConnectors.Schema,
  ): Promise<void> {
    const response: Response = await this.graphClient
      .api(`/external/connections/${connectionId}/schema`)
      .responseType(ResponseType.RAW)
      .post(schema);

    if (response.ok) {
      // The operation ID is contained in the Location header returned
      // in the response
      const location = response.headers.get('Location');
      const locationSegments = location?.split('/') ?? [];
      if (locationSegments.length <= 0) {
        throw new Error('Could not get operation ID from Location header');
      }

      const operationId = locationSegments[locationSegments.length - 1];
      await this.waitForOperationToCompleteAsync(connectionId, operationId);
    } else {
      throw new Error(
        `Registering schema failed, status: ${response.status} - ${response.statusText}`,
      );
    }
  }

  /**
   * Adds or updates an ExternalItem.
   *
   * @param connectionId - The connection ID of the connection that contains the item.
   * @param item - The item to add or update.
   * @returns The item.
   */
  public async addOrUpdateItemAsync(
    connectionId: string,
    item: ExternalConnectors.ExternalItem,
  ): Promise<ExternalConnectors.ExternalItem> {
    return this.graphClient
      .api(`/external/connections/${connectionId}/items/${item.id}`)
      .put(item);
  }

  public async addIssueActivitiesAsync(
    connectionId: string,
    itemId: string,
    activities: ExternalActivityWithType[],
  ) {
    if (activities.length > 0) {
      await this.graphClient
        .api(
          `/external/connections/${connectionId}/items/${itemId}/addActivities`,
        )
        .post({
          activities: activities,
        });
    }
  }

  /**
   * Creates an ExternalItem from a Repository.
   *
   * @param repo - The repository.
   * @param repoEvents - A list of repository events, used to determine the use that last modified the repo.
   * @returns The ExternalItem.
   */
  public async createExternalItemFromRepoAsync(
    repo: Repository,
    repoEvents: RepoEvent[],
  ): Promise<ExternalConnectors.ExternalItem> {
    let lastModifiedBy = repo.owner.login;
    if (repoEvents && repoEvents.length > 0) {
      lastModifiedBy = repoEvents[repoEvents.length - 1].actor.login;
    }

    const activities: ExternalActivityWithType[] = [
      {
        '@odata.type': externalActivityType,
        type: 'created',
        startDateTime: repo.created_at ?? undefined,
        performedBy: await this.getIdentityForGitHubUserAsync(repo.owner.login),
      },
    ];

    const externalItem: ExternalConnectors.ExternalItem = {
      id: repo.id.toString(),
      acl: [
        {
          type: 'everyone',
          value: 'everyone',
          accessType: 'grant',
        },
      ],
      properties: {
        title: repo.name,
        description: repo.description,
        visibility: repo.visibility ?? 'unknown',
        createdBy: repo.owner.login,
        updatedAt: repo.updated_at,
        lastModifiedBy: lastModifiedBy,
        repoUrl: repo.html_url,
        userUrl: repo.owner.html_url,
        icon: 'https://pngimg.com/uploads/github/github_PNG40.png',
      },
      activities: activities,
    };

    return externalItem;
  }

  /**
   * Creates an ExternalItem from an Issue.
   * @param issue - The issue.
   * @param issueEvents - A list of issue events, used to determine the use that last modified the issue.
   * @returns The ExternalItem.
   */
  public async createExternalItemFromIssueAsync(
    issue: Issue,
    issueEvents: IssueEvent[],
  ): Promise<ExternalConnectors.ExternalItem> {
    let lastModifiedBy = issue.user?.login;
    if (issueEvents && issueEvents.length > 0) {
      lastModifiedBy =
        issueEvents[issueEvents.length - 1].actor?.login ?? lastModifiedBy;
    }

    const activities: ExternalActivityWithType[] = [
      {
        '@odata.type': externalActivityType,
        type: 'created',
        startDateTime: issue.created_at ?? undefined,
        performedBy: await this.getIdentityForGitHubUserAsync(
          issue.user?.login,
        ),
      },
    ];

    const externalItem: ExternalConnectors.ExternalItem = {
      id: issue.number.toString(),
      acl: [
        {
          type: 'everyone',
          value: 'everyone',
          accessType: 'grant',
        },
      ],
      properties: {
        title: issue.title,
        body: issue.body,
        assignees: this.assigneesToString(issue.assignees ?? []),
        labels: this.labelsToString(issue.labels),
        state: issue.state,
        issueUrl: issue.html_url,
        lastModifiedBy: lastModifiedBy,
        updatedAt: issue.updated_at,
        icon: 'https://pngimg.com/uploads/github/github_PNG40.png',
      },
      activities: activities,
    };

    return externalItem;
  }

  /**
   * Creates a list of ExternalActivityWithTypes from a list of IssueEvents.
   * @param issueEvents - The list of IssueEvents.
   * @returns The list of ExternalActivityWithTypes.
   */
  public async createExternalActivitiesFromIssueEventsAsync(
    issueEvents: IssueEvent[],
  ): Promise<ExternalActivityWithType[]> {
    const activities: ExternalActivityWithType[] = [];

    for (const issueEvent of issueEvents) {
      activities.push({
        '@odata.type': externalActivityType,
        type: issueEvent.event === 'commented' ? 'commented' : 'modified',
        startDateTime: issueEvent.created_at,
        performedBy: await this.getIdentityForGitHubUserAsync(
          issueEvent.actor?.login,
        ),
      });
    }
    return activities;
  }

  /**
   * Periodically polls a newly created operation to check for completion.
   * @param connectionId - The connection ID of the connection.
   * @param operationId - The operation ID of the operation to check.
   */
  private async waitForOperationToCompleteAsync(
    connectionId: string,
    operationId: string,
  ) {
    let keepPolling = true;
    do {
      const operation = (await this.graphClient
        .api(`/external/connections/${connectionId}/operations/${operationId}`)
        .get()) as ExternalConnectors.ConnectionOperation;

      if (operation.status === 'completed') {
        keepPolling = false;
      } else if (operation.status == 'failed') {
        throw new Error(
          operation.error?.message ?? 'Registering schema failed',
        );
      } else {
        // Poll every minute
        await new Promise((res) => setTimeout(res, 60000));
      }
    } while (keepPolling);
  }

  /**
   * Gets an Identity from a GitHub login.
   * @param _ - The GitHub login to look up.
   * @returns The Identity.
   */
  private async getIdentityForGitHubUserAsync(
    _?: string,
  ): Promise<ExternalConnectors.Identity> {
    // This method simply returns the user ID set in
    // .env. A possible improvement would be
    // to have some mapping of GitHub logins to Azure AD users

    return {
      type: 'user',
      id: this.placeHolderUserId,
    };
  }

  /**
   * Converts a list of assignees to a string.
   * @param assignees - the list of assignees.
   * @returns A comma-delimited string of GitHub logins.
   */
  private assigneesToString(assignees: Assignee[]): string {
    if (assignees.length <= 0) {
      return 'None';
    }

    return assignees.map((a) => a.login).join(',');
  }

  /**
   * Converts a list of labels to a string.
   * @param labels - the list of labels.
   * @returns A comma-delimited string of label names.
   */
  private labelsToString(labels?: Labels): string {
    if (!labels || labels.length <= 0) {
      return 'None';
    }

    // Labels can be a plain string or a label object
    return labels.map((l) => (typeof l === 'string' ? l : l.name)).join(',');
  }

  /**
   * Loads adaptive card layout from a file.
   * @param resultCardJsonFile - the path to the file.
   * @returns The parsed JSON layout.
   */
  private getResultTemplate(resultCardJsonFile: string) {
    const resultTemplate = readFileSync(resultCardJsonFile, 'utf-8');
    return JSON.parse(resultTemplate);
  }
}
