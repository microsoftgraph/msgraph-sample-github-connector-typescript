// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export enum MenuChoice {
  // Exit
  Exit = -1,
  // Create a new connection
  CreateConnection,
  // Select an existing connection
  SelectConnection,
  // Delete the current connection
  DeleteConnection,
  // Register schema on the current connection
  RegisterSchema,
  // Push items to the current connection
  PushAllItems,
  // Invalid choice
  Invalid,
}

export const menuPrompts = [
  'Create a connection',
  'Select existing connection',
  'Delete current connection',
  'Register schema for current connection',
  'Push items to current connection',
];

export enum ItemTypeChoice {
  Issues,
  Repositories,
}

export const itemTypes = ['Issues', 'Repositories'];
