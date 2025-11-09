/**
 * @OnlyCurrentDoc
 *
 * Global constants for the LARP Downtime Management Script.
 */

// ============================================================================
// === CONSTANTS ============================================================
// ============================================================================

/**
 * Keywords used to categorize downtime actions. The key is the category,
 * and the value is an array of strings that match that category (case-insensitive).
 * @const {!Object<string, !Array<string>>}
 */
const DOWNTIME_KEYWORDS = {
  feed: ['feed'],
  patrol: ['patrol'],
  investigate: ['investigate'],
  'beyond your means': ['beyond your means'],
  quest: ['quest'],
  'eternal struggle': ['eternal struggle', 'elder game', 'elder downtime'],
  'learn disciplines': ['learn', 'learn disciplines'],
  rhetorical: ['rhetorical'],
  narrative: ['narrative'],
  observe: ['observe', 'spy'],
  cancel: ['cancel', 'block'],
};

/**
 * Definitions for Discord markdown formatting buttons used in the editor.
 * @const {!Array<{name: string, prefix: string, suffix: string}>}
 */
const DISCORD_MARKDOWN_STYLES = [
  { name: 'Bold', prefix: '**', suffix: '**' },
  { name: 'Italics', prefix: '_', suffix: '_' },
  { name: 'Underline', prefix: '__', suffix: '__' },
  { name: 'Strikethrough', prefix: '~~', suffix: '~~' },
  { name: 'Spoiler', prefix: '||', suffix: '||' },
  { name: 'Code', prefix: '`', suffix: '`' },
  { name: 'Code Block', prefix: '```\n', suffix: '\n```' },
];

/** Regular expression to identify downtime-related columns in the header row. */
const DOWNTIME_HEADER_REGEX = /downtime/i;
/** Regular expression to identify influence-related columns in the header row. */
const INFLUENCE_HEADER_REGEX = /influence/i;
/** Regular expression to identify resource-related columns in the header row. */
const RESOURCES_HEADER_REGEX = /resources/i;
/** Regular expression to match sheet names like "Month Year". */
const MONTH_YEAR_SHEET_REGEX = /^\w+ \d{4}$/; // Matches "MonthName 2024" etc.


/** Column index (1-based) for the Timestamp in the downtime sheet. */
const TIMESTAMP_COL = 1;
/** Column index (1-based) for the Status in the downtime sheet. */
const STATUS_COL = 2;
/** Column index (1-based) for the "Send Discord" checkbox in the downtime sheet. */
const SEND_DISCORD_COL = 3;
/** Column index (1-based) for the "Send Email" checkbox in the downtime sheet. */
const SEND_EMAIL_COL = 4;
/** Column index (1-based) for the Character Name in the downtime sheet. */
const CHARACTER_NAME_COL = 5;
/** Column index (1-based) for the Character Name in the "Characters" sheet. */
const CHAR_SHEET_NAME_COL = 1;
/** Column index (1-based) for the Discord Webhook in the "Characters" sheet. */
const CHAR_SHEET_WEBHOOK_COL = 24; // Column X
/** Column index (1-based) for the Approval Status in the "Characters" sheet. */
const CHAR_SHEET_APPROVAL_COL = 8; // Column H

/** Name of the sheet containing character data. */
const CHARACTER_SHEET_NAME = 'Characters';
/** Name of the sheet containing NPC data. */
const NPC_SHEET_NAME = 'NPCs';
/** Name of the sheet containing feed list data. */
const FILL_REPLACE_SHEET_NAME = 'fill-replace-data';
/** Name of the sheet listing character influences. */
const INFLUENCES_SHEET_NAME = 'Influences';
/** Name of the sheet listing elite influence details. */
const ELITE_INFLUENCES_SHEET_NAME = 'Elite Infl.';
/** Name of the sheet listing underworld influence details. */
const UW_INFLUENCES_SHEET_NAME = 'UW Infl.';
/** Name of the sheet for logging actions. */
const AUDIT_LOG_SHEET_NAME = 'Log';
/** Name of the sheet for scheduled messages. */
const SCHEDULED_MSG_SHEET_NAME = 'Scheduled Messages'; // Added constant

/** Maximum length for a single Discord message chunk. */
const MAX_DISCORD_MESSAGE_LENGTH = 1800; // Slightly less than 2000 for safety

/** Script Property keys */
const PROP_ST_WEBHOOK = 'ST_WEBHOOK';
const PROP_DOWNTIME_FORM_ID = 'DOWNTIME_FORM_ID';
const PROP_DOWNTIME_YEAR = 'DOWNTIME_YEAR';
const PROP_DOWNTIME_MONTH = 'DOWNTIME_MONTH';
const PROP_PDF_FOLDER_ID = 'PDF_FOLDER_ID';
const PROP_DISCORD_TEST_MODE = 'DISCORD_TEST_MODE';
const PROP_TEST_WEBHOOK = 'TEST_WEBHOOK';
const PROP_ANNOUNCEMENT_WEBHOOK = 'ANNOUNCEMENT_WEBHOOK'; // For #announcements channel
const PROP_IC_CHAT_WEBHOOK = 'IC_CHAT_WEBHOOK';         // *** NEW *** For #ic-chat channel
const PROP_IC_NEWS_FEED_WEBHOOK = 'IC_NEWS_FEED_WEBHOOK'; // *** NEW *** For #ic-news-feed channel
const PROP_LARP_NAME = 'LARP_NAME';
const PROP_TASK_COLOR_HEX = 'TASK_COLOR_HEX'; // For 'Any Narrator or Storyteller' assignment

/** Discord Send Retry Constants */
const MAX_RETRIES = 2; // Max number of retries (total 3 attempts)
const INITIAL_BACKOFF_MS = 1000; // Initial delay in milliseconds for exponential backoff
const MAX_BACKOFF_MS = 60000; // Maximum delay capped at 60 seconds

/** Fallback text for influence fill */
const INFLUENCE_FALLBACK_TEXT = "Either nothing happened or tracks were sufficiently covered.";

/** Color Constants for Validation */
const COLOR_VALID = '#CCFFCC'; // Light Green
const COLOR_NO_WEBHOOK = '#FFCCCC'; // Light Red
const COLOR_NO_MATCH = '#FFDAB9'; // PeachPuff (Light Orange)

/** Scheduled Messages Sheet Column Indices (1-based) */
const SCHEDULED_MSG_SENT_COL = 1;      // Column A: Sent (Checkbox)
const SCHEDULED_MSG_TIME_COL = 2;      // Column B: Send After Timestamp
const SCHEDULED_MSG_CHANNEL_COL = 3;   // Column C: Target Channel (Dropdown)
const SCHEDULED_MSG_SENDER_COL = 4;    // *** NEW *** Column D: Sender (Optional)
const SCHEDULED_MSG_BODY_COL = 5;      // Column E: Message Body (Shifted)
const SCHEDULED_MSG_LOG_COL = 6;       // Column F: Send Log (Shifted)

/** Characters Sheet Column Indices (1-based) */
// const CHAR_SHEET_NAME_COL = 1; // Already defined above
// const CHAR_SHEET_APPROVAL_COL = 8; // Already defined above
// const CHAR_SHEET_WEBHOOK_COL = 24; // Already defined above
const NPC_SHEET_AVATAR_COL_SHEET_AVATAR_COL = 25; // *** NEW *** Column Y: Avatar URL

/** Master list of all possible Influence Specializations */
const ELITE_INFL_SHEET_NAME = 'Elite Infl.';
const UW_INFL_SHEET_NAME = 'UW Infl.';
const ALL_SPECIALIZATIONS = [
    "Academics", "Activist Organizations", "Arts", "Commercial Industry",
    "Criminal Organizations", "Emergency Services", "Entertainment", "Finance",
    "High Society", "Infrastructure", "Manufacturing", "Law Enforcement",
    "Legal", "Local Government", "Science/Medical", "Social Media",
    "News Media", "Occult", "Religious Groups/Organizations", "Transportation"
].sort(); // Sort master list alphabetically

/** Influence Summary Dialog Constants */
const TOP_N_SPECS = 10; // How many items to show in Top/Bottom lists

/** Influence Fill Action Column Indices (0-based) for Elite/UW Sheets */
// Assumes format: Date(A), Age(B), Character(C), Specialization(D), Action(E), Details(F), Blocks?(G), Output(H)
const INFL_DATE_COL_IDX = 0; // Column A
const INFL_SPEC_COL_IDX = 3; // Column D
const INFL_ACTION_COL_IDX = 4; // Column E
