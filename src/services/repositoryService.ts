// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Octokit } from '@octokit/rest';
import { retry } from '@octokit/plugin-retry';
import { components } from '@octokit/openapi-types';

const OctokitWithRetry = Octokit.plugin(retry);

// Load schemas from GitHub's OpenAPI types
export type Repository = components['schemas']['repository'];
export type Readme = components['schemas']['content-file'];
export type RepoEvent = components['schemas']['event'];
export type Issue = components['schemas']['issue'];
export type IssueComment = components['schemas']['issue-comment'];
export type IssueEvent = components['schemas']['issue-event'];
export type Assignee = components['schemas']['simple-user'];
export type Labels = components['schemas']['issue']['labels'];

export type RepositoryServiceOptions = {
  gitHubToken?: string;
  gitHubOwner?: string;
  gitHubRepo?: string;
};

export default class RepositoryService {
  private gitHubClient: Octokit;
  private gitHubOwner: string;
  private gitHubRepo: string;

  constructor(options: RepositoryServiceOptions) {
    if (!options.gitHubToken || !options.gitHubOwner || !options.gitHubRepo) {
      throw new Error('Invalid GitHub details, please see README');
    }
    this.gitHubClient = new OctokitWithRetry({
      auth: options.gitHubToken,
    });
    this.gitHubOwner = options.gitHubOwner;
    this.gitHubRepo = options.gitHubRepo;
  }

  /**
   * Gets a list of repositories for the user or organization specified in app settings.
   *
   * @returns The list of repositories.
   */
  public async getRepositoriesAsync(): Promise<Repository[] | undefined> {
    try {
      // Assume owner is an organization
      return (await this.gitHubClient.paginate('GET /orgs/{org}/repos', {
        org: this.gitHubOwner,
      })) as Repository[];
    } catch (error) {
      if ((error as Error).message === 'Not Found') {
        // If not found as an organization, try as a user
        return (await this.gitHubClient.paginate(
          'GET /users/{username}/repos',
          {
            username: this.gitHubOwner,
          },
        )) as Repository[];
      } else {
        throw error;
      }
    }
  }

  /**
   * Gets the README for a repository.
   *
   * @param repoName - The name of the repository.
   * @returns The README.
   */
  public async getReadmeAsync(repoName: string): Promise<Readme> {
    const response = await this.gitHubClient.request(
      'GET /repos/{owner}/{repo}/readme',
      {
        owner: this.gitHubOwner,
        repo: repoName,
      },
    );

    return response.data as Readme;
  }

  /**
   * Gets activity events for a repository.
   *
   * @param repoName - The name of the repository.
   * @returns The list of events.
   */
  public async getEventsForRepoAsync(repoName: string) {
    return (await this.gitHubClient.paginate(
      'GET /repos/{owner}/{repo}/events',
      {
        owner: this.gitHubOwner,
        repo: repoName,
      },
    )) as RepoEvent[];
  }

  /**
   * Gets all issues for the GitHub repository specified in app settings.
   * @returns The list of issues.
   */
  public async getIssuesForRepositoryAsync(): Promise<Issue[]> {
    return (await this.gitHubClient.paginate(
      'GET /repos/{owner}/{repo}/issues',
      {
        owner: this.gitHubOwner,
        repo: this.gitHubRepo,
      },
    )) as Issue[];
  }

  public async getCommentsForIssueAsync(
    issueNumber: number,
  ): Promise<IssueComment[]> {
    return (await this.gitHubClient.paginate(
      'GET /repos/{owner}/{repo}/issues/{issue_number}/comments',
      {
        owner: this.gitHubOwner,
        repo: this.gitHubRepo,
        issue_number: issueNumber,
      },
    )) as IssueComment[];
  }

  /**
   * Gets all events for a GitHub issue.
   * @param issueNumber - The issue number.
   * @returns The list of events.
   */
  public async getEventsForIssueAsync(
    issueNumber: number,
  ): Promise<IssueEvent[]> {
    return (await this.gitHubClient.paginate(
      'GET /repos/{owner}/{repo}/issues/{issue_number}/events',
      {
        owner: this.gitHubOwner,
        repo: this.gitHubRepo,
        issue_number: issueNumber,
      },
    )) as IssueEvent[];
  }
}
