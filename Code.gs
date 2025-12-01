/**
 * MEETING NOTES → LINEAR AUTOMATION
 * FINAL CLEAN VERSION — Bookmark placed directly on the “Action Items” heading.
 * Script regenerates all content BELOW that heading.
 */

// ==================== CONFIGURATION ====================
function getConfig() {
  const props = PropertiesService.getScriptProperties().getProperties();

  return {
    LINEAR_API_KEY: props.LINEAR_API_KEY,
    LINEAR_TEAM_ID: props.LINEAR_TEAM_ID,
    LINEAR_PROJECT_ID: props.LINEAR_PROJECT_ID,
    ACTION_ITEMS_BOOKMARK: props.ACTION_ITEMS_BOOKMARK, // required bookmark id
    GEMINI_API_KEY: props.GEMINI_API_KEY // required for AI extraction
  };
}
// =====================================================

const CONFIG = getConfig();
let LINEAR_USERS = {};
let MEETING_NOTES_LABEL_ID = null;

/**
 * Adds custom menu when document opens
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Action Items')
    .addItem('Open Action Sidebar', 'showSidebar')
    .addItem('Go to Action Items Tab', 'jumpToActionItems')
    .addToUi();
}

/**
 * Opens sidebar
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Meeting Actions');
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Jumps to the user-created bookmark at the heading
 */
function jumpToActionItems() {
  const doc = DocumentApp.getActiveDocument();
  const bookmarks = doc.getBookmarks();

  const bookmark = bookmarks.find(b => b.getId() === CONFIG.ACTION_ITEMS_BOOKMARK);
  if (!bookmark) return `Bookmark ${CONFIG.ACTION_ITEMS_BOOKMARK} not found.`;

  doc.setCursor(bookmark.getPosition());
  return `Jumped to Action Items section`;
}

/**
 * BUTTON 1: Update Action Items
 * Clears all content below the bookmarked heading and regenerates it in-place.
 */
function updateActionItemsSection() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const bookmark = doc.getBookmarks().find(b => b.getId() === CONFIG.ACTION_ITEMS_BOOKMARK);

  if (!bookmark) {
    return `Bookmark ${CONFIG.ACTION_ITEMS_BOOKMARK} not found.`;
  }

  const pos = bookmark.getPosition();
  const headingElement = pos.getElement();

  // Ensure the bookmark is on a heading paragraph
  const heading = headingElement.getType() === DocumentApp.ElementType.PARAGRAPH
    ? headingElement.asParagraph()
    : headingElement.getParent().asParagraph();

  const startIndex = body.getChildIndex(heading);
  clearSectionBelow(body, startIndex);

  // Generate new content below heading
  let insertIndex = startIndex + 1;
  const actions = getOpenActionsFromLast4WeeksInThisDoc();

  body.insertParagraph(insertIndex++, `Updated: ${new Date().toLocaleString()}`).setItalic(true);
  body.insertHorizontalRule(insertIndex++);

  if (actions.length === 0) {
    body.insertParagraph(insertIndex++, 'No open action items from the last 4 weeks.');
  } else {
    // Create checklist
    const first = body.insertListItem(insertIndex++, actions[0].text);
    first.setGlyphType(DocumentApp.GlyphType.SQUARE);
    const listId = first.getListId();

    for (let i = 1; i < actions.length; i++) {
      const item = body.insertListItem(insertIndex++, actions[i].text);
      item.setGlyphType(DocumentApp.GlyphType.SQUARE);
      item.setListId(listId);
    }
  }

  return `Updated ${actions.length} action item(s).`;
}

/**
 * Removes everything below the heading until the next heading or end of document
 */
function clearSectionBelow(body, startIndex) {
  let i = startIndex + 1;

  while (i < body.getNumChildren()) {
    const child = body.getChild(i);

    // Stop at next H1 heading
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH &&
        child.asParagraph().getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
      break;
    }

    child.removeFromParent();
  }
}

/**
 * Scans full document and extracts open action items from last 4 weeks using Gemini AI
 */
function getOpenActionsFromLast4WeeksInThisDoc() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const fullText = body.getText();
  
  // Get last 4 weeks of content
  const fourWeeksAgo = new Date(new Date().setDate(new Date().getDate() - 28));
  const formattedDate = fourWeeksAgo.toISOString().split('T')[0];
  
  // Call Gemini API to extract action items
  const prompt = `You are analyzing meeting notes to extract action items. 
  
Extract all action items from meetings dated ${formattedDate} or later from the following text.

For each action item, identify:
1. The assignee name (person responsible)
2. The action description

Return ONLY a JSON array of objects with this exact format:
[
  {"assignee": "FirstName LastName", "description": "Action item text"},
  {"assignee": "FirstName LastName", "description": "Action item text"}
]

Rules:
- Only include incomplete/open action items
- Do NOT include items already marked as done or completed
- Keep descriptions concise but complete
- If no assignee is clear, use "Unassigned"
- Only return the JSON array, no other text

Meeting notes text:

${fullText}`;

  const actions = callGeminiAPI(prompt);
  
  // Convert to the format expected by the rest of the script
  const actionsMap = {};
  for (const action of actions) {
    const key = `${action.description.toLowerCase()}|||${action.assignee}`;
    if (!actionsMap[key]) {
      actionsMap[key] = {
        text: action.assignee === 'Unassigned' 
          ? action.description 
          : `@${action.assignee} ${action.description}`
      };
    }
  }
  
  return Object.values(actionsMap);
}

/**
 * Calls Gemini API to analyze text and extract action items
 */
function callGeminiAPI(prompt) {
  if (!CONFIG.GEMINI_API_KEY) {
    throw new Error('GEMINI_API_KEY not configured in Script Properties');
  }
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
  
  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 2048
    }
  };
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const json = JSON.parse(response.getContentText());
    
    if (json.error) {
      throw new Error(`Gemini API Error: ${json.error.message}`);
    }
    
    const text = json.candidates[0].content.parts[0].text;
    
    // Extract JSON from the response (handle markdown code blocks if present)
    const jsonMatch = text.match(/\[[\s\S]*\]/);
    if (!jsonMatch) {
      Logger.log('Gemini response: ' + text);
      return [];
    }
    
    return JSON.parse(jsonMatch[0]);
    
  } catch (e) {
    Logger.log('Error calling Gemini API: ' + e.message);
    throw new Error('Failed to extract action items using AI: ' + e.message);
  }
}

/**
 * BUTTON 2: Push to Linear
 */
function pushActionsToLinear() {
  const actions = extractOpenItemsFromGeneratedSection();
  if (actions.length === 0) return 'No open actions to push.';

  LINEAR_USERS = fetchLinearUsers();
  MEETING_NOTES_LABEL_ID = ensureMeetingNotesLabel();

  let pushed = 0;

  for (const a of actions) {
    if (a.hasLinearId) continue;

    const issueId = createLinearIssue(a);
    if (issueId) {
      a.element.setText(`${a.cleanText} (${issueId})`);
      pushed++;
    }
  }

  return pushed > 0
    ? `Pushed ${pushed} action(s) to Linear.`
    : 'All items already pushed.';
}

/**
 * Reads generated section only (using heading marker)
 */
function extractOpenItemsFromGeneratedSection() {
  const paragraphs = DocumentApp.getActiveDocument().getBody().getParagraphs();
  const result = [];
  let inSection = false;

  for (const p of paragraphs) {
    const text = p.getText();

    if (text.includes('Action Items (last 4 weeks)') ||
        text.includes('Action Items')) {
      inSection = true;
      continue;
    }

    if (!inSection) continue;
    if (p.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

    const li = p.asListItem();
    const isUnchecked = [
      DocumentApp.GlyphType.SQUARE,
      DocumentApp.GlyphType.HOLLOW_BULLET
    ].includes(li.getGlyphType());

    if (!isUnchecked) continue;

    const full = text.replace(/^[-•·]\s*/, '').trim();
    const linearMatch = full.match(/\((\w+-\d+)\)$/);
    const clean = full.replace(/\(\w+-\d+\)$/, '').trim();

    const assignee = clean.match(/@(\w+)/)?.[1] || 'Unassigned';
    const description = clean.replace(/^@\w+\s+/, '');

    result.push({
      element: p,
      cleanText: clean,
      description,
      assignee,
      hasLinearId: !!linearMatch
    });
  }

  return result;
}

/**
 * Creates Linear issue
 */
function createLinearIssue(action) {
  const assigneeId = LINEAR_USERS[action.assignee.toLowerCase()] || null;

  const query = `
    mutation CreateIssue($input: IssueCreateInput!) {
      issueCreate(input: $input) {
        success
        issue { identifier }
      }
    }`;

  const variables = {
    input: {
      teamId: CONFIG.LINEAR_TEAM_ID,
      projectId: CONFIG.LINEAR_PROJECT_ID,
      title: action.description.substring(0, 200),
      description: `**Assignee**: @${action.assignee}
**Source**: ${DocumentApp.getActiveDocument().getUrl()}`,
      labelIds: MEETING_NOTES_LABEL_ID ? [MEETING_NOTES_LABEL_ID] : [],
      assigneeId: assigneeId
    }
  };

  const resp = UrlFetchApp.fetch('https://api.linear.app/graphql', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: CONFIG.LINEAR_API_KEY },
    payload: JSON.stringify({ query, variables }),
    muteHttpExceptions: true
  });

  const json = JSON.parse(resp.getContentText());
  return json.data?.issueCreate?.success
    ? json.data.issueCreate.issue.identifier
    : null;
}

/**
 * Fetch Linear users
 */
function fetchLinearUsers() {
  const resp = UrlFetchApp.fetch('https://api.linear.app/graphql', {
    method: 'post',
    headers: { Authorization: CONFIG.LINEAR_API_KEY },
    payload: JSON.stringify({ query: '{ users { id name } }' })
  });

  const users = JSON.parse(resp.getContentText()).data.users;
  const map = {};

  for (const u of users) {
    map[u.name.split(' ')[0].toLowerCase()] = u.id;
  }
  return map;
}

/**
 * Ensure label exists
 */
function ensureMeetingNotesLabel() {
  const resp = UrlFetchApp.fetch('https://api.linear.app/graphql', {
    method: 'post',
    headers: { Authorization: CONFIG.LINEAR_API_KEY },
    payload: JSON.stringify({ query: '{ labels { nodes { id name } } }' })
  });

  const labels = JSON.parse(resp.getContentText()).data.labels.nodes;
  const existing = labels.find(l => l.name.toLowerCase() === 'meeting notes');
  if (existing) return existing.id;

  const create = UrlFetchApp.fetch('https://api.linear.app/graphql', {
    method: 'post',
    headers: { Authorization: CONFIG.LINEAR_API_KEY },
    payload: JSON.stringify({
      query: 'mutation { labelCreate(input: {name: "Meeting Notes", color: "#4f46e5"}) { label { id } } }'
    })
  });

  return JSON.parse(create.getContentText()).data.labelCreate.label.id;
}
