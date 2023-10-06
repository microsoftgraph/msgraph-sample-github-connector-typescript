// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ExternalConnectors } from '@microsoft/microsoft-graph-types-beta';

/**
 * Extends the ExternalConnectors.ExternalActivity to give access to
 * the \@odata.type property.
 *
 * @remarks
 * This is needed because Microsoft Graph will return a 400 InvalidRequest
 * error if the \@odata.type property isn't present in the POST body
 * to create a new activity.
 */
export default interface ExternalActivityWithType
  extends ExternalConnectors.ExternalActivity {
  '@odata.type': string;
}

export const externalActivityType =
  '#microsoft.graph.externalConnectors.externalActivity';
