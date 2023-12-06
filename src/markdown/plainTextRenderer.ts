// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { MarkedOptions, Renderer } from 'marked';

/**
 * Custom renderer for marked.js to convert
 * Markdown to plain text
 */
export default class PlainTextRenderer extends Renderer {
  constructor(options?: MarkedOptions) {
    super(options);
  }

  code(code: string, _info: string | undefined, _escaped: boolean): string {
    return '\n\n' + code + '\n\n';
  }

  blockquote(quote: string): string {
    return quote + '\n';
  }

  html(html: string, _block?: boolean): string {
    return html;
  }

  heading(text: string, _level: number, _raw: string): string {
    return text;
  }

  hr(): string {
    return '\n\n';
  }

  list(body: string, _ordered: boolean, _start: number | ''): string {
    return body;
  }

  listitem(text: string, _task: boolean, _checked: boolean): string {
    return '- ' + text + '\n';
  }

  checkbox(_checked: boolean): string {
    return '';
  }

  paragraph(text: string): string {
    return '\n' + text + '\n';
  }

  table(header: string, body: string): string {
    return '\n' + header + '\n' + body + '\n';
  }

  tablerow(content: string): string {
    return content + '\n';
  }

  tablecell(
    content: string,
    _flags: { header: boolean; align: 'center' | 'left' | 'right' | null },
  ): string {
    return content + '\t';
  }

  strong(text: string): string {
    return text;
  }

  em(text: string): string {
    return text;
  }

  codespan(text: string): string {
    return text;
  }

  br(): string {
    return '\n\n';
  }

  del(text: string): string {
    return text;
  }

  link(_href: string, _title: string | null | undefined, text: string): string {
    return text;
  }

  image(_href: string, _title: string | null, text: string): string {
    return text;
  }

  text(text: string): string {
    return text;
  }
}
