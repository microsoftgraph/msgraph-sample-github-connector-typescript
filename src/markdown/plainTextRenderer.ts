// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { MarkedOptions, Renderer, Tokens } from 'marked';

/**
 * Custom renderer for marked.js to convert
 * Markdown to plain text
 */
export default class PlainTextRenderer extends Renderer {
  constructor(options?: MarkedOptions) {
    super(options);
  }

  code(tokens: Tokens.Code): string {
    return '\n\n' + tokens.text + '\n\n';
  }

  blockquote(tokens: Tokens.Blockquote): string {
    return tokens.text + '\n';
  }

  html(tokens: Tokens.HTML): string {
    return tokens.text;
  }

  heading(tokens: Tokens.Heading): string {
    return tokens.text;
  }

  hr(): string {
    return '\n\n';
  }

  list(_tokens: Tokens.List): string {
    return '';
  }

  listitem(tokens: Tokens.ListItem): string {
    return '- ' + tokens.text + '\n';
  }

  checkbox(_tokens: Tokens.Checkbox): string {
    return '';
  }

  paragraph(tokens: Tokens.Paragraph): string {
    return '\n' + tokens.text + '\n';
  }

  table(_tokens: Tokens.Table): string {
    return '';
  }

  tablerow(tokens: Tokens.TableRow): string {
    return tokens.text + '\n';
  }

  tablecell(tokens: Tokens.TableCell): string {
    return tokens.text + '\t';
  }

  strong(tokens: Tokens.Strong): string {
    return tokens.text;
  }

  em(tokens: Tokens.Em): string {
    return tokens.text;
  }

  codespan(tokens: Tokens.Codespan): string {
    return tokens.text;
  }

  br(): string {
    return '\n\n';
  }

  del(tokens: Tokens.Del): string {
    return tokens.text;
  }

  link(tokens: Tokens.Link): string {
    return tokens.text + ' (' + tokens.href + ')';
  }

  image(tokens: Tokens.Image): string {
    return tokens.text;
  }

  text(tokens: Tokens.Text): string {
    return tokens.text;
  }
}
