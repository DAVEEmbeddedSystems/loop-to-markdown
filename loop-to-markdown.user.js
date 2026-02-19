// ==UserScript==
// @name         Microsoft Loop to Markdown
// @namespace    http://tampermonkey.net/
// @version      1.6
// @description  Convert Microsoft Loop pages to Markdown
// @author       Talha Oz (ozt@), Yuta TJ (yutatj@), Shrinivas Acharya
// @match        https://loop.cloud.microsoft/*
// @match        https://*.loop.cloud.microsoft.com/*
// @grant        GM_setClipboard
// @updateURL    https://raw.githubusercontent.com/oztalha/loop-to-markdown/main/loop-to-markdown.user.js
// @downloadURL  https://raw.githubusercontent.com/oztalha/loop-to-markdown/main/loop-to-markdown.user.js
// @license      GPL-3.0
// ==/UserScript==

/*
 * Changelog:
 * v1.6 - Merged contributions:
 *   yutatj@: getTextContentFallback() for table cells with only tags, table header fallback
 *   Shrinivas: Bold text detection, code language auto-detection, ordered list support,
 *              correct heading levels (aria-level), code duplicate prevention
 * v1.5 - yutatj@: Fixed empty table cells with links/mentions
 * v1.4 - ozt@: Initial release - DOM parsing, tables, code blocks, headings, lists,
 *              checkboxes, @mentions, hyperlinks, inline code, quip link capture
 */

(function() {
    'use strict';

    const normalize = text => {
        if (!text) return '';
        let result = text.trim().replace(/\s+/g, ' ');
        // Fix empty/malformed bold markers
        result = result.replace(/\*\*\*\*/g, '').replace(/\*\*\s*\*\*/g, '');
        result = result.replace(/(\w)\*\*(?=\w)/g, '$1 **');
        result = result.replace(/(\S)\*\*(\w)/g, '$1** $2');
        return result.trim();
    };

    const getMention = el => {
        const avatar = el.querySelector('.fui-Avatar[aria-label]');
        return avatar ? `@${avatar.getAttribute('aria-label')}` : '';
    };

    const getTextContent = (container, skipTables = false) => {
        let text = '';
        const targets = container.querySelectorAll('.scriptor-textRun, [data-testid="resolvedAtMention"]');
        if (targets.length === 0) {
            return getTextContentFallback(container);
        }
        targets.forEach(node => {
            if (skipTables && node.closest('table')) return;
            if (node.dataset.testid === 'resolvedAtMention') {
                text += getMention(node);
            } else if (!node.closest('[data-testid="resolvedAtMention"]')) {
                if (node.classList.contains('scriptor-hyperlink')) {
                    const href = (node.getAttribute('title') || '').split('\n')[0];
                    if (href) text += `[${node.textContent}](${href})`;
                } else if (node.classList.contains('scriptor-code-editor')) {
                    text += '`' + node.textContent + '`';
                } else {
                    let content = node.textContent;
                    const style = window.getComputedStyle(node);
                    const isBold = parseInt(style.fontWeight) >= 600 || style.fontWeight === 'bold' || style.fontWeight === 'bolder';
                    if (isBold && content.trim()) {
                        const lead = content.match(/^\s*/)[0], trail = content.match(/\s*$/)[0];
                        text += `${lead}**${content.trim()}**${trail}`;
                    } else {
                        text += content;
                    }
                }
            }
        });
        return normalize(text);
    };

    const getTextContentFallback = (container) => {
        let text = '';
        const walk = (node) => {
            if (node.nodeType === Node.TEXT_NODE) { text += node.textContent; return; }
            if (node.nodeType !== Node.ELEMENT_NODE) return;
            const el = node, tag = el.tagName.toLowerCase();
            if (el.dataset?.testid === 'resolvedAtMention') {
                const mention = getMention(el);
                if (mention) { text += mention; return; }
            }
            if (tag === 'a') {
                const href = el.getAttribute('href') || el.getAttribute('title')?.split('\n')[0] || '';
                const linkText = el.textContent.trim();
                if (href && linkText) { text += `[${linkText}](${href})`; return; }
            }
            if (el.classList?.contains('scriptor-code-editor')) { text += '`' + el.textContent + '`'; return; }
            for (const child of el.childNodes) walk(child);
        };
        for (const child of container.childNodes) walk(child);
        return normalize(text);
    };

    const detectCodeLanguage = (code, element) => {
        const langAttr = element?.getAttribute('data-language') ||
            element?.closest('[data-language]')?.getAttribute('data-language') ||
            element?.querySelector('[data-language]')?.getAttribute('data-language');
        if (langAttr) return langAttr.toLowerCase();
        const trimmed = code.trim();
        if (/^(flowchart|graph|sequenceDiagram|classDiagram|stateDiagram|erDiagram|gantt|pie|journey)\s/i.test(trimmed)) return 'mermaid';
        if (/^(def |class |import |from |async def |@\w+)/.test(trimmed)) return 'python';
        if (/^(const |let |var |function |import |export |async |=>|interface |type |enum )/.test(trimmed)) return 'javascript';
        if (/^(public |private |protected |class |interface |package |import java)/.test(trimmed)) return 'java';
        if (/^(SELECT|INSERT|UPDATE|DELETE|CREATE|ALTER|DROP)\s/i.test(trimmed)) return 'sql';
        if (/^\w+:\s*(\n|$)/.test(trimmed) && !trimmed.includes('{') && trimmed.includes(':')) return 'yaml';
        if (/^[\[{]/.test(trimmed) && /[\]}]$/.test(trimmed)) return 'json';
        if (/^(#!\/|npm |yarn |pip |git |docker |kubectl |curl |wget |\$ )/.test(trimmed)) return 'bash';
        if (/^<(!DOCTYPE|html|head|body|div|span|p|a|script)/i.test(trimmed)) return 'html';
        if (/^(\.|#|@media|@import|\*|body|html)\s*{/.test(trimmed)) return 'css';
        return '';
    };

    const parseTable = table => {
        const lines = [], headers = [];
        table.querySelectorAll('[role="columnheader"]').forEach(th => {
            const label = th.querySelector('[aria-label]');
            headers.push(label ? label.getAttribute('aria-label') : getTextContent(th) || '');
        });
        if (headers.length) {
            lines.push('| ' + headers.join(' | ') + ' |', '| ' + headers.map(() => '---').join(' | ') + ' |');
        }
        table.querySelectorAll('tbody tr[data-rowid]').forEach(row => {
            if (row.dataset.rowid === 'HEADER_ROW_ID') return;
            const cells = [...row.querySelectorAll('[role="cell"]')].map(cell => getTextContent(cell).replace(/\|/g, '\\|'));
            if (cells.length) lines.push('| ' + cells.join(' | ') + ' |');
        });
        return lines;
    };

    function convertToMarkdown() {
        const pages = [...document.querySelectorAll('.scriptor-pageFrame')].filter(p => !p.closest('table'));
        if (!pages.length) return alert('No Loop content found');

        const lines = [], processed = new Set(), codeTexts = new Set(), codeRawTexts = new Set();

        const firstPara = pages[0]?.querySelector('.scriptor-paragraph:not([role="heading"] *)');
        if (firstPara && !firstPara.querySelector('[role="heading"]') && !firstPara.closest('.scriptor-listItem, table')) {
            const title = normalize(firstPara.textContent);
            if (title) { lines.push(`# ${title}`, ''); processed.add(firstPara); }
        }

        pages.forEach(page => {
            page.querySelectorAll('.scriptor-paragraph, .scriptor-listItem, .scriptor-component-code-block, [role="table"], [role="heading"]').forEach(el => {
                if (processed.has(el)) return;
                const inTable = el.closest('table');
                if (inTable && inTable !== el) return;

                if (el.getAttribute('role') === 'table') {
                    const tableLines = parseTable(el);
                    if (tableLines.length) lines.push('', ...tableLines, '');
                    processed.add(el);
                    return;
                }

                if (el.classList.contains('scriptor-paragraph') && el.closest('.scriptor-component-code-block')) return;
                const codeBlock = el.querySelector('.scriptor-code-wrap-on') ||
                    (el.classList.contains('scriptor-component-code-block') ? el.querySelector('.scriptor-code-editor') : null);
                if (codeBlock) {
                    const code = [...codeBlock.querySelectorAll('.scriptor-paragraph')].map(p => p.textContent).join('\n').trim() || codeBlock.textContent.trim();
                    if (code) {
                        const lang = detectCodeLanguage(code, el);
                        lines.push('', '```' + lang, code, '```', '');
                        codeTexts.add(normalize(code));
                        codeRawTexts.add(code.replace(/\s+/g, ' ').trim());
                        codeBlock.querySelectorAll('.scriptor-paragraph').forEach(p => processed.add(p));
                        processed.add(el);
                    }
                    return;
                }

                const heading = el.getAttribute('role') === 'heading' ? el : el.querySelector('[role="heading"]');
                if (heading) {
                    const level = parseInt(heading.getAttribute('aria-level') || '1', 10);
                    let text = getTextContent(heading, true).replace(/\*\*/g, '').trim();
                    if (text) lines.push('', `${'#'.repeat(Math.min(level, 6))} ${text}`, '');
                    processed.add(el);
                    return;
                }

                if (el.classList.contains('scriptor-listItem')) {
                    const li = el.querySelector('li');
                    if (!li) return;
                    const text = getTextContent(li);
                    if (!text) return;
                    const margin = parseInt((el.getAttribute('style') || '').match(/margin-left:\s*(\d+)/)?.[1] || 0);
                    const indent = '  '.repeat(Math.max(0, Math.floor((margin - 27) / 27)));
                    const checkbox = li.querySelector('.scriptor-listItem-marker-checkbox');
                    const checked = checkbox?.getAttribute('aria-checked') === 'true';
                    const listParent = li.closest('ol, ul');
                    const markerEl = el.querySelector('.scriptor-listItem-marker, [class*="listItem-marker"]');
                    const markerText = markerEl?.textContent?.trim() || '';
                    const hasNumberMarker = /^\d+[\.\)]?$/.test(markerText);
                    const dataListType = el.getAttribute('data-list-type') || el.closest('[data-list-type]')?.getAttribute('data-list-type');
                    const isOrdered = listParent?.tagName === 'OL' || hasNumberMarker || dataListType === 'ordered' || dataListType === 'number';
                    let marker;
                    if (checkbox) {
                        marker = checked ? '- [x] ' : '- [ ] ';
                    } else if (isOrdered) {
                        const numMatch = markerText.match(/^(\d+)/);
                        const value = numMatch ? numMatch[1] : (li.getAttribute('value') || '1');
                        marker = `${value}. `;
                    } else {
                        marker = '- ';
                    }
                    lines.push(indent + marker + text);
                    processed.add(el);
                    return;
                }

                if (!el.closest('.scriptor-listItem') && !el.querySelector('.scriptor-code-wrap-on')) {
                    let text = getTextContent(el, true);
                    el.querySelectorAll('a[href*="quip"]').forEach(link => {
                        const href = link.getAttribute('href');
                        if (href) text += ` [${link.textContent.trim() || href.split('/').pop()}](${href})`;
                    });
                    text = normalize(text);
                    const isCodeDuplicate = [...codeTexts].some(c => c.includes(text) || text.includes(c)) ||
                        [...codeRawTexts].some(c => c.includes(text) || text.includes(c));
                    if (text && !isCodeDuplicate) lines.push('', text, '');
                    processed.add(el);
                }
            });
        });

        let markdown = lines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
        markdown = markdown.replace(/(```\w*\n[\s\S]*?\n```)\n\n\1/g, '$1');
        GM_setClipboard(markdown, 'text');

        const note = document.createElement('div');
        note.textContent = '✓ Markdown copied!';
        note.style.cssText = 'position:fixed;top:20px;right:20px;background:#4CAF50;color:white;padding:12px 16px;border-radius:5px;z-index:10000;font-family:sans-serif';
        document.body.appendChild(note);
        setTimeout(() => note.remove(), 2000);
    }

    const btn = document.createElement('button');
    btn.textContent = '📋 Copy as Markdown';
    btn.style.cssText = 'position:fixed;bottom:20px;left:20px;background:#0078D4;color:white;border:none;padding:8px 12px;border-radius:5px;cursor:pointer;z-index:10000;font-family:sans-serif;font-size:12px';
    btn.onclick = convertToMarkdown;
    document.body.appendChild(btn);
})();
