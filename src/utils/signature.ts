import { existsSync, readFileSync } from 'fs';
import { homedir } from 'os';
import { join } from 'path';

const CONFIG_DIR = join(homedir(), '.config', 'myoffice-mcp');

/**
 * Awolve has two signatures: a standard one with logo and full contact details
 * for new mail, and a minimal single line for replies. `none` sends neither.
 */
export type SignatureStyle = 'standard' | 'minimal' | 'none';

/**
 * A signature has an HTML rendering and may have a plain-text one beside it
 * (`signature.txt`, `signature-minimal.txt`). Only the HTML files are installed
 * today; the text ones are read when present so a text/plain body never gets
 * markup pasted into it (myoffice bug #2).
 */
export type SignatureFormat = 'html' | 'text';

const FILES: Record<SignatureFormat, Record<Exclude<SignatureStyle, 'none'>, string>> = {
  html: {
    standard: join(CONFIG_DIR, 'signature.html'),
    minimal: join(CONFIG_DIR, 'signature-minimal.html'),
  },
  text: {
    standard: join(CONFIG_DIR, 'signature.txt'),
    minimal: join(CONFIG_DIR, 'signature-minimal.txt'),
  },
};

export function isSignatureStyle(value: unknown): value is SignatureStyle {
  return value === 'standard' || value === 'minimal' || value === 'none';
}

/**
 * Read a signature. Returns null when the style is `none` or the file has not
 * been installed — in Cortex, `bun scripts/cortex-signature.ts` writes both from
 * the person's own context, and /cortex-doctor reports when they are missing.
 *
 * A missing minimal signature falls back to the standard one rather than
 * sending nothing: an over-formal sign-off beats a reply that looks anonymous.
 */
export function getSignature(
  style: SignatureStyle = 'standard',
  format: SignatureFormat = 'html'
): string | null {
  if (style === 'none') return null;
  const files = FILES[format];
  const candidates = style === 'minimal' ? [files.minimal, files.standard] : [files.standard];
  for (const path of candidates) {
    try {
      if (!existsSync(path)) continue;
      const content = readFileSync(path, 'utf-8').trim();
      if (content) return content;
    } catch {
      // unreadable — try the next candidate, then give up quietly
    }
  }
  return null;
}

/**
 * Append the chosen signature to a body, or return the body untouched if there
 * is none. A plain-text body only ever gets the plain-text rendering: when no
 * `.txt` signature is installed, nothing is appended rather than the HTML one.
 */
export function appendSignature(body: string, isHtml: boolean, style: SignatureStyle): string {
  const signature = getSignature(style, isHtml ? 'html' : 'text');
  if (!signature) return body;
  return isHtml ? `${body}<br><br>${signature}` : `${body}\n\n--\n${signature}`;
}
