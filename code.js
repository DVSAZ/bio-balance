/*************************************************************
 *  GLOBAL CONSTANTS
 *************************************************************/
const SUBMISSIONS_SPREADSHEET_ID = "1gv_qdttwlZb8-ifHsKBuw1v_mr2XLwDkttGzDXk5tG4";
const SUBMISSIONS_SHEET_NAME = "Submissions";
const CUSTOMERS_SHEET_NAME = "Submissions";
const HEADER_ROW_INDEX = 1;
const CLOSE_AUDIT_SHEET_NAME = "Close Audit";
const FOLLOW_UP_PROPERTY_PREFIX = "FOLLOWUP_";
const TRIGGER_LOOKUP_PREFIX = "TRIGGER_";
const GA4_MEASUREMENT_ID_PROP = "GA4_MEASUREMENT_ID";
const GA4_API_SECRET_PROP = "GA4_API_SECRET";
const META_PIXEL_ID_PROP = "META_PIXEL_ID";
const META_ACCESS_TOKEN_PROP = "META_ACCESS_TOKEN";
const CLOSE_EMAIL_FROM_PROP = "CLOSE_EMAIL_FROM";
const FOLLOW_UP_EMAIL_FROM_PROP = "FOLLOW_UP_EMAIL_FROM";
const REPORT_STATUS_SHEET_NAME = "ReportStatus";
const REPORT_STATUS_HEADERS = ["leadID", "stepNumber", "statusText", "lastUpdated"];

/*************************************************************
 *  CLEAN VALUES TO PREVENT ROW EXPANSION
 *************************************************************/
function cleanValue(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value;
  const str = String(value).replace(/\n/g, " ").trim();
  return str;
}

function applyQuizNotesToRow(rowValues, quizNotes) {
  if (!Array.isArray(quizNotes) || !rowValues) return;

  quizNotes.forEach(function(note) {
    if (!note || typeof note !== "object") return;
    const key = note.key || note.fieldKey;
    if (!key || !QUIZ_FIELD_COLUMN_MAP[key]) return;
    const columnIndex = QUIZ_FIELD_COLUMN_MAP[key] - 1;
    const value = cleanValue(note.note || note.value || note.text || "");
    if (!value) return;
    rowValues[columnIndex] = value;
  });
}

/*************************************************************
 *  COLUMN INDEX CONSTANTS (1-based)
 *  Updated to match header:
 *  Timestamp | Lead ID | Source | UTM Campaign | UTM Ad Set | UTM Keyword | UTM Landing Page | Gender | Name | Email | Phone | Age Range | Sleep Quality | Energy Levels | Mental Clarity | Stress Tolerance | Physical Activity | Body Changes | Libido | Temperature Sensitivity (F) | Cycle Status (F) | PMS / Mood Shifts (F) | Bloating / Fluid Retention (F) | Strength & Recovery (M) | Morning Drive / Motivation (M) | Muscle Mass / Strength Changes (M) | Sexual Performance / Libido (M) | Bio-Balance Score (%) | Lead Score (Sales) | Promo Code | Promo Description | Promo Redeemed? | B12 Offered? | B12 Redeemed? | Lead Status | Assigned Staff | Follow-Up Priority | Last Outreach Date | Next Follow-Up Date | Contact Attempts | Last Attempt Method | Last Attempt Result | Opted Out? | Notes | Booked Consultation? | Consultation Date | Joined Membership? | Membership Type | Membership Start Date | Cold Lead Retargeting Bucket | Tags | Raw Answers (JSON) | Lead Type | Follow-Up Sent? | 72hr Follow-Up Timestamp | Closed From
 *************************************************************/
const COL = {
  TIMESTAMP: 1,
  LEAD_ID: 2,
  SOURCE: 3,
  UTM_CAMPAIGN: 4,
  UTM_AD_SET: 5,
  UTM_KEYWORD: 6,
  UTM_LANDING_PAGE: 7,
  GENDER: 8,
  NAME: 9,
  EMAIL: 10,
  PHONE: 11,
  AGE_RANGE: 12,
  SLEEP_QUALITY: 13,
  ENERGY_LEVELS: 14,
  MENTAL_CLARITY: 15,
  STRESS_TOLERANCE: 16,
  PHYSICAL_ACTIVITY: 17,
  BODY_CHANGES: 18,
  LIBIDO: 19,
  TEMP_SENSITIVITY_F: 20,
  CYCLE_STATUS_F: 21,
  PMS_MOOD_F: 22,
  BLOATING_F: 23,
  STRENGTH_RECOVERY_M: 24,
  MORNING_DRIVE_M: 25,
  MUSCLE_MASS_M: 26,
  SEXUAL_PERFORMANCE_M: 27,
  BIO_BALANCE_SCORE: 28,
  LEAD_SCORE_SALES: 29,
  PROMO_CODE: 30,
  PROMO_DESCRIPTION: 31,
  PROMO_REDEEMED: 32,
  B12_OFFERED: 33,
  B12_REDEEMED: 34,
  LEAD_STATUS: 35,
  ASSIGNED_STAFF: 36,
  FOLLOW_UP_PRIORITY: 37,
  LAST_OUTREACH_DATE: 38,
  NEXT_FOLLOW_UP_DATE: 39,
  CONTACT_ATTEMPTS: 40,
  LAST_ATTEMPT_METHOD: 41,
  LAST_ATTEMPT_RESULT: 42,
  OPTED_OUT: 43,
  NOTES: 44,
  BOOKED_CONSULTATION: 45,
  CONSULTATION_DATE: 46,
  JOINED_MEMBERSHIP: 47,
  MEMBERSHIP_TYPE: 48,
  MEMBERSHIP_START_DATE: 49,
  RETARGET_BUCKET: 50,
  TAGS: 51,
  RAW_ANSWERS: 52,
  LEAD_TYPE: 53,
  FOLLOW_UP_SENT: 54,
  FOLLOW_UP_72HR_TIMESTAMP: 55,
  CLOSED_FROM: 56
};

const QUIZ_FIELD_COLUMN_MAP = {
  gender: COL.GENDER,
  ageRange: COL.AGE_RANGE,
  sleepQuality: COL.SLEEP_QUALITY,
  energyLevels: COL.ENERGY_LEVELS,
  mentalClarity: COL.MENTAL_CLARITY,
  stressTolerance: COL.STRESS_TOLERANCE,
  activityLevel: COL.PHYSICAL_ACTIVITY,
  bodyChanges: COL.BODY_CHANGES,
  libido: COL.LIBIDO,
  tempSensitivityF: COL.TEMP_SENSITIVITY_F,
  cycleStatusF: COL.CYCLE_STATUS_F,
  pmsMoodF: COL.PMS_MOOD_F,
  bloatingF: COL.BLOATING_F,
  strengthRecoveryM: COL.STRENGTH_RECOVERY_M,
  morningDriveM: COL.MORNING_DRIVE_M,
  muscleMassM: COL.MUSCLE_MASS_M,
  sexualPerformanceM: COL.SEXUAL_PERFORMANCE_M
};

const SHARED_SYMPTOM_FIELD_DEFS = [
  { key: "sleepQuality", label: "Sleep" },
  { key: "energyLevels", label: "Energy" },
  { key: "mentalClarity", label: "Mental Clarity" },
  { key: "stressTolerance", label: "Stress" },
  { key: "activityLevel", label: "Activity" },
  { key: "bodyChanges", label: "Body Changes" },
  { key: "libido", label: "Libido" }
];

const FEMALE_SYMPTOM_FIELD_DEFS = [
  { key: "tempSensitivityF", label: "Temp Sensitivity" },
  { key: "cycleStatusF", label: "Cycle Status" },
  { key: "pmsMoodF", label: "PMS / Mood" },
  { key: "bloatingF", label: "Bloating" }
];

const MALE_SYMPTOM_FIELD_DEFS = [
  { key: "strengthRecoveryM", label: "Strength & Recovery" },
  { key: "morningDriveM", label: "Morning Drive" },
  { key: "muscleMassM", label: "Muscle Mass" },
  { key: "sexualPerformanceM", label: "Sexual Performance" }
];

const PIPELINE_STAGE_KEYS = ["quizComplete", "bookedConsult", "consultNoJoin", "member"];

const SYMPTOM_FIELD_LOOKUP = (function() {
  const map = {};
  [SHARED_SYMPTOM_FIELD_DEFS, FEMALE_SYMPTOM_FIELD_DEFS, MALE_SYMPTOM_FIELD_DEFS].forEach(function(group) {
    group.forEach(function(field) {
      const keyNormalized = normalizeFilterString(field.key);
      if (keyNormalized && !map[keyNormalized]) {
        map[keyNormalized] = field.key;
      }
      const labelNormalized = normalizeFilterString(field.label);
      if (labelNormalized && !map[labelNormalized]) {
        map[labelNormalized] = field.key;
      }
    });
  });
  return map;
})();

const SYMPTOM_FIELD_THRESHOLDS = {
  default: 2,
  sexualPerformanceM: 3
};

const SYMPTOM_TEXT_SEVERITY_RULES = {
  sleepQuality: [
    { score: 3, pattern: /regular problem/i },
    { score: 2, pattern: /restless or interrupted/i },
    { score: 1, pattern: /occasional restless/i }
  ],
  energyLevels: [
    { score: 3, pattern: /running on low battery/i },
    { score: 2, pattern: /feel tired more often/i },
    { score: 1, pattern: /comes and goes/i }
  ],
  mentalClarity: [
    { score: 3, pattern: /sluggish|slow frequently/i },
    { score: 2, pattern: /lose focus|feel foggy/i },
    { score: 1, pattern: /mild fog|forgetfulness/i }
  ],
  stressTolerance: [
    { score: 3, pattern: /overwhelming|hard to recover/i },
    { score: 2, pattern: /drains me faster/i },
    { score: 1, pattern: /i manage/i }
  ],
  activityLevel: [
    { score: 3, pattern: /mostly inactive|too tired/i },
    { score: 2, pattern: /consistency is tough/i },
    { score: 1, pattern: /exercise a few times/i }
  ],
  bodyChanges: [
    { score: 3, pattern: /changing even though/i },
    { score: 2, pattern: /maintaining my weight/i },
    { score: 1, pattern: /minor changes/i }
  ],
  libido: [
    { score: 3, pattern: /very low|almost nonexistent/i },
    { score: 2, pattern: /noticeably decreased/i },
    { score: 1, pattern: /slightly lower/i }
  ],
  tempSensitivityF: [
    { score: 3, pattern: /hard to ignore/i },
    { score: 2, pattern: /more frequently/i },
    { score: 1, pattern: /occasionally/i }
  ],
  cycleStatusF: [
    { score: 3, pattern: /post-menopausal/i },
    { score: 2, pattern: /perimenopause/i },
    { score: 1, pattern: /irregular/i }
  ],
  pmsMoodF: [
    { score: 3, pattern: /significant changes/i },
    { score: 2, pattern: /more noticeable/i },
    { score: 1, pattern: /mild changes/i }
  ],
  bloatingF: [
    { score: 3, pattern: /most days/i },
    { score: 2, pattern: /often/i },
    { score: 1, pattern: /sometimes/i }
  ],
  strengthRecoveryM: [
    { score: 3, pattern: /major decrease/i },
    { score: 2, pattern: /slower recovery|declining strength/i },
    { score: 1, pattern: /slight slowdown/i }
  ],
  morningDriveM: [
    { score: 3, pattern: /rare lately/i },
    { score: 2, pattern: /low-energy/i },
    { score: 1, pattern: /some mornings/i }
  ],
  muscleMassM: [
    { score: 3, pattern: /significant decrease/i },
    { score: 2, pattern: /noticeable loss/i },
    { score: 1, pattern: /slight decrease/i }
  ],
  sexualPerformanceM: [
    { score: 4, pattern: /noticeably decreased/i },
    { score: 3, pattern: /occasional difficulty/i },
    { score: 2, pattern: /libido is lower/i }
  ]
};

function getSymptomThreshold(fieldKey) {
  if (Object.prototype.hasOwnProperty.call(SYMPTOM_FIELD_THRESHOLDS, fieldKey)) {
    return SYMPTOM_FIELD_THRESHOLDS[fieldKey];
  }
  return SYMPTOM_FIELD_THRESHOLDS.default;
}

function parseLeadRawAnswers(rawAnswers) {
  if (!rawAnswers) {
    return null;
  }
  if (typeof rawAnswers === "object") {
    return rawAnswers;
  }
  if (typeof rawAnswers !== "string") {
    return null;
  }
  const trimmed = rawAnswers.trim();
  if (!trimmed) {
    return null;
  }
  try {
    return JSON.parse(trimmed);
  } catch (err) {
    return null;
  }
}

function computeSeverityFromRawEntry(rawEntry) {
  if (rawEntry === null || rawEntry === undefined) {
    return null;
  }
  if (typeof rawEntry === "number") {
    return isNaN(rawEntry) ? null : rawEntry;
  }
  if (Array.isArray(rawEntry)) {
    return rawEntry.reduce(function(max, value) {
      const num = Number(value);
      if (!isNaN(num) && num > max) {
        return num;
      }
      return max;
    }, 0);
  }
  if (typeof rawEntry === "object") {
    if (Array.isArray(rawEntry.values)) {
      return rawEntry.values.reduce(function(max, value) {
        const num = Number(value);
        if (!isNaN(num) && num > max) {
          return num;
        }
        return max;
      }, 0);
    }
    if (rawEntry.value !== undefined) {
      const num = Number(rawEntry.value);
      return isNaN(num) ? 0 : num;
    }
  }
  if (typeof rawEntry === "string") {
    const num = Number(rawEntry);
    if (!isNaN(num)) {
      return num;
    }
  }
  return null;
}

function inferSymptomSeverityFromText(fieldKey, value) {
  const normalized = normalizeTagValue(value);
  if (!normalized) {
    return 0;
  }
  const rules = SYMPTOM_TEXT_SEVERITY_RULES[fieldKey];
  if (!rules || !rules.length) {
    return 0;
  }
  const parts = normalized.toLowerCase().split(/\s*,\s*/);
  let severity = 0;
  parts.forEach(function(part) {
    rules.forEach(function(rule) {
      if (rule.pattern.test(part)) {
        if (rule.score > severity) {
          severity = rule.score;
        }
      }
    });
  });
  return severity;
}

function getSymptomSeverityForLead(lead, rawAnswers, fieldKey) {
  const rawEntry = rawAnswers && rawAnswers[fieldKey];
  let severity = computeSeverityFromRawEntry(rawEntry);
  if (severity === null || severity === undefined) {
    severity = inferSymptomSeverityFromText(fieldKey, lead[fieldKey]);
  }
  if (!severity || !isFinite(severity) || severity < 0) {
    return 0;
  }
  return severity;
}

function normalizeTagValue(tag) {
  if (tag === null || tag === undefined) {
    return "";
  }
  const trimmed = String(tag).trim();
  if (!trimmed) {
    return "";
  }
  return trimmed.replace(/\s+/g, " ");
}

function normalizeTagForComparison(tag) {
  const normalized = normalizeTagValue(tag);
  return normalized ? normalized.toLowerCase() : "";
}

function normalizeFilterString(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim().toLowerCase();
}

function cleanTagList(tags) {
  if (!Array.isArray(tags)) {
    return [];
  }
  const cleaned = [];
  const seen = new Set();
  tags.forEach(function(tag) {
    const normalized = normalizeTagValue(tag);
    if (!normalized) {
      return;
    }
    const key = normalized.toLowerCase();
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    cleaned.push(normalized);
  });
  return cleaned;
}

function mergeAutoAndManualTags(autoTags, manualTags) {
  const merged = [];
  const seen = new Set();

  function push(tag) {
    const normalized = normalizeTagValue(tag);
    if (!normalized) {
      return;
    }
    // Skip removal markers in final output
    if (normalized.startsWith('!REMOVE:')) {
      return;
    }
    const key = normalized.toLowerCase();
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    merged.push(normalized);
  }

  (autoTags || []).forEach(push);
  (manualTags || []).forEach(push);

  return merged;
}

function deriveManualTags(allTags, autoTags) {
  const autoSet = new Set((autoTags || []).map(normalizeTagForComparison));
  if (!Array.isArray(allTags)) {
    return [];
  }
  const manual = allTags.filter(function(tag) {
    return !autoSet.has(normalizeTagForComparison(tag));
  });
  return cleanTagList(manual);
}

function parseTagsInput(input) {
  if (Array.isArray(input)) {
    return cleanTagList(input);
  }
  if (input === null || input === undefined) {
    return [];
  }
  const normalized = normalizeTagValue(input);
  if (!normalized) {
    return [];
  }
  const parts = normalized.split(/[;,]+/);
  if (parts.length === 1) {
    return cleanTagList([normalized]);
  }
  return cleanTagList(parts);
}

function isMeaningfulTagValue(value) {
  const normalized = normalizeTagValue(value);
  if (!normalized) {
    return false;
  }
  const lowered = normalized.toLowerCase();
  if (lowered === "none" || lowered === "n/a" || lowered === "na" || lowered === "normal" || lowered === "unknown" || lowered === "-" || lowered === "—" || lowered === "--") {
    return false;
  }
  return true;
}

function formatDateForTag(value) {
  if (!value) {
    return "";
  }
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "MMM d, yyyy");
  }
  const normalized = normalizeTagValue(value);
  if (!normalized) {
    return "";
  }
  const parsed = new Date(normalized);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "MMM d, yyyy");
  }
  return normalized;
}

function computeAutoTagsFromLead(lead) {
  if (!lead) {
    return [];
  }

  const autoTags = [];
  const seen = new Set();

  // Get manual tags to check for removal markers
  const manualTags = Array.isArray(lead.manualTags) ? lead.manualTags : parseTagsInput(lead.manualTags || "");
  const removedTags = new Set();
  manualTags.forEach(function(tag) {
    const normalized = normalizeTagValue(tag);
    if (normalized && normalized.startsWith('!REMOVE:')) {
      const tagToRemove = normalized.substring(8); // Remove '!REMOVE:' prefix
      removedTags.add(normalizeTagForComparison(tagToRemove));
    }
  });

  function add(tag) {
    const normalized = normalizeTagValue(tag);
    if (!normalized) {
      return;
    }
    const key = normalized.toLowerCase();
    if (seen.has(key) || removedTags.has(key)) {
      return;
    }
    seen.add(key);
    autoTags.push(normalized);
  }

  // Auto-tag: Gender
  const gender = normalizeTagValue(lead.gender || lead.genderIdentity);
  if (gender) {
    add("Gender: " + gender);
  }

  // Auto-tag: Age or Age Range
  const ageValue = normalizeTagValue(lead.ageRange || lead.ageBracket || lead.age);
  if (ageValue) {
    add("Age: " + ageValue);
  }

  // Auto-tag: Membership Type (if they joined)
  if (lead.joinedMembership) {
    const membershipType = normalizeTagValue(lead.membershipType || lead.membershipPlan);
    if (membershipType) {
      add("Membership: " + membershipType);
    }
  }

  // Auto-tag: Promo Offers
  const promoOffered = normalizeTagValue(lead.promoDescription || lead.promoDetails || lead.promoCode);
  if (promoOffered) {
    add("Promo: " + promoOffered);
  }

  // Auto-tag: Provider (if assigned)
  const assignedStaff = normalizeTagValue(lead.assignedStaff || lead.provider || lead.assignedProvider);
  if (assignedStaff) {
    // Normalize provider names
    const staffLower = assignedStaff.toLowerCase();
    if (staffLower.includes("amie") || staffLower.includes("amy")) {
      add("Provider: Amie");
    } else if (staffLower.includes("amber")) {
      add("Provider: Amber");
    } else if (staffLower.includes("laura")) {
      add("Provider: Laura");
    } else {
      // Add as-is if it's a provider name we don't recognize
      add("Provider: " + assignedStaff);
    }
  }

  return autoTags;
}

/*************************************************************
 *  UNIVERSAL CORS HANDLER
 *************************************************************/
function sendCorsResponse(content) {
  const output = ContentService
    .createTextOutput(content)
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

/*************************************************************
 *  REQUIRED FOR BROWSER PREFLIGHT REQUESTS
 *************************************************************/
function doOptions(e) {
  return sendCorsResponse(JSON.stringify({ status: "ok" }));
}

/*************************************************************
 *  OPTIONAL: Handle GET (just returns a safe JSON notice)
 *************************************************************/
function doGet(e) {
  try {
    const leadId = (e?.parameter?.leadID || e?.parameter?.leadId || "").toString().trim();

    if (!leadId) {
      return sendCorsResponse(JSON.stringify({
        step: 0,
        statusText: "Missing leadID"
      }));
    }

    const status = getStatusByLeadId(leadId);
    return sendCorsResponse(JSON.stringify(status));
  } catch (err) {
    return sendCorsResponse(JSON.stringify({
      step: 0,
      statusText: "error",
      message: err.toString()
    }));
  }
}

/*************************************************************
 *  UPDATE LEAD FIELDS FROM CALL CENTER WEBHOOK
 *************************************************************/
function updateLeadFromCall(payload) {
  if (!payload || payload.rowNumber === undefined || payload.rowNumber === null) {
    throw new Error("Missing rowNumber in payload.");
  }

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Lead Center");
  if (!sheet) {
    throw new Error("Sheet 'Lead Center' not found. Update sheet name if needed.");
  }

  const row = Number(payload.rowNumber);
  if (!row || row < 2) {
    throw new Error("Invalid rowNumber provided.");
  }

  const LEAD_STATUS_COL = 32;        // Update if your sheet uses a different column
  const FOLLOW_UP_PRIORITY_COL = 34; // Update if your sheet uses a different column
  const TAGS_COL = 55;               // Update if your sheet uses a different column

  const updatedFields = [];

  if (payload.leadStatus !== undefined) {
    sheet.getRange(row, LEAD_STATUS_COL).setValue(payload.leadStatus);
    updatedFields.push("leadStatus");
  }

  if (payload.followUpPriority !== undefined) {
    sheet.getRange(row, FOLLOW_UP_PRIORITY_COL).setValue(payload.followUpPriority);
    updatedFields.push("followUpPriority");
  }

  if (payload.tags !== undefined) {
    const tagsValue = Array.isArray(payload.tags) ? payload.tags.join(",") : payload.tags;
    sheet.getRange(row, TAGS_COL).setValue(tagsValue);
    updatedFields.push("tags");
  }

  return {
    success: true,
    updatedFields: updatedFields
  };
}

/*************************************************************
 *  HELPER: Ensure Gender_Log Sheet Exists
 *************************************************************/
function ensureGenderSheet() {
  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  let sh = ss.getSheetByName("Gender_Log");
  if (!sh) {
    sh = ss.insertSheet("Gender_Log");
    sh.appendRow(["Timestamp", "Gender", "ClientTimestamp"]);
  }
}

/*************************************************************
 *  HELPER: Ensure Internal_Events Sheet Exists
 *************************************************************/
function ensureInternalEventsSheet() {
  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  let sh = ss.getSheetByName("Internal_Events");
  if (!sh) {
    sh = ss.insertSheet("Internal_Events");
    sh.appendRow(["Timestamp", "Event", "Data", "ClientTimestamp"]);
  }
}

/*************************************************************
 *  MAIN: Handle Webflow / Live Server POST Request
 *************************************************************/
function doPost(e) {
  ensureGenderSheet();
  ensureInternalEventsSheet();
  let leadId = "";
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No POST data received.");
    }

    const payload = JSON.parse(e.postData.contents);
    const action = (payload.action || "").toString().toLowerCase();

    if (action === "close_report" || action === "reportclose" || action === "close") {
      const closeResult = processReportClose(payload);
      return sendCorsResponse(JSON.stringify(closeResult));
    }

    if (action === "sendreminderemail") {
      const reminderResult = sendFollowUpReminder(payload || {});
      return sendCorsResponse(JSON.stringify(reminderResult));
    }

    if (action === "updateleadfromcall") {
      const updateResult = updateLeadFromCall(payload || {});
      return sendCorsResponse(JSON.stringify(updateResult));
    }

    if (payload.type === "gender_selection") {
      const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
      const sh = ss.getSheetByName("Gender_Log");

      sh.appendRow([
        new Date(),
        payload.gender,
        payload.timestamp
      ]);

      return ContentService
        .createTextOutput(JSON.stringify({status: "logged"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.type === "internal_event") {
      const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
      const sh = ss.getSheetByName("Internal_Events");

      sh.appendRow([
        new Date(),
        payload.event,
        JSON.stringify(payload.data),
        payload.timestamp
      ]);

      return ContentService
        .createTextOutput(JSON.stringify({ status: "event_logged" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);
    if (!sheet) throw new Error("Sheet 'Submissions' not found.");

    const now = new Date();

    const normalizeField = function(value) {
      if (value === null || value === undefined) return "";
      if (Array.isArray(value)) return value.join(", ");
      if (typeof value === "object") {
        if (Array.isArray(value.texts)) return value.texts.map(String).join(", ");
        if (typeof value.text === "string") return value.text;
        return "";
      }
      return String(value);
    };

    const gender   = payload.gender || "";
    const name     = payload.name || payload.contactInfo?.name || "";
    const email    = payload.email || payload.contactInfo?.email || "";
    const phone    = payload.phone || payload.contactInfo?.phone || "";
    const ageRange = normalizeField(payload.ageRange);

    const sleepQuality    = normalizeField(payload.sleepQuality);
    const energyLevels    = normalizeField(payload.energyLevels);
    const mentalClarity   = normalizeField(payload.mentalClarity);
    const stressTolerance = normalizeField(payload.stressTolerance);
    const activityLevel   = normalizeField(payload.physicalActivity || payload.activityLevel);
    const bodyChanges     = normalizeField(payload.bodyChanges);
    const libido          = normalizeField(payload.libido);

    const genderLower = gender.toLowerCase();
    const tempSensitivityF = genderLower === "female" ? normalizeField(payload.tempSensitivityF) : "";
    const cycleStatusF     = genderLower === "female" ? normalizeField(payload.cycleStatusF) : "";
    const pmsMoodF         = genderLower === "female" ? normalizeField(payload.pmsMoodF) : "";
    const bloatingF        = genderLower === "female" ? normalizeField(payload.bloatingF) : "";

    const strengthRecoveryM  = genderLower === "male" ? normalizeField(payload.strengthRecoveryM) : "";
    const morningDriveM      = genderLower === "male" ? normalizeField(payload.morningDriveM) : "";
    const muscleMassM        = genderLower === "male" ? normalizeField(payload.muscleMassM) : "";
    const sexualPerformanceM = genderLower === "male" ? normalizeField(payload.sexualPerformanceM) : "";

    const utmCampaign    = payload.utmCampaign || payload.utm_campaign || "";
    const utmAdSet       = payload.utmAdSet || payload.utm_adset || "";
    const utmKeyword     = payload.utmKeyword || payload.utm_keyword || "";
    const utmLandingPage = payload.utmLandingPage || payload.utm_landing_page || "";

    const source = payload.source || "Bio-Balance Landing Page";
    const rawAnswers = typeof payload.rawAnswersJson === "string"
      ? payload.rawAnswersJson.replace(/\n/g, " ")
      : (typeof payload.rawAnswers === "string"
        ? payload.rawAnswers.replace(/\n/g, " ")
        : JSON.stringify(payload).replace(/\n/g, " "));

    const leadType = payload.leadType || "Quiz Lead";
    const promoCode = payload.promoCode || "";
    const promoDescription = payload.promoDescription || autoPromoDescription(promoCode);

    // Generate Lead ID
    leadId = payload.leadId || generateLeadId(sheet.getLastRow() + 1);
    updateStatus(leadId, 1, "Collecting your responses");

    // Calculate initial lead score
    const scorePercent = Number(payload.bioBalanceScore ?? payload.scorePercent ?? 0) || 0;
    const symptomData = {
      ageRange, sleepQuality, energyLevels, mentalClarity,
      stressTolerance, email, phone, name
    };
    const leadScoreSales = payload.leadScoreSales !== undefined && payload.leadScoreSales !== ""
      ? payload.leadScoreSales
      : calculateSalesLeadScore(scorePercent, symptomData);

    /*************************************************************
     *  SAVE TO GOOGLE SHEET - Strict header order
     *************************************************************/
    const row = Array(COL.CLOSED_FROM).fill("");
    row[COL.TIMESTAMP - 1] = now;
    row[COL.LEAD_ID - 1] = leadId;
    row[COL.SOURCE - 1] = source;
    row[COL.UTM_CAMPAIGN - 1] = utmCampaign;
    row[COL.UTM_AD_SET - 1] = utmAdSet;
    row[COL.UTM_KEYWORD - 1] = utmKeyword;
    row[COL.UTM_LANDING_PAGE - 1] = utmLandingPage;
    row[COL.GENDER - 1] = gender;
    row[COL.NAME - 1] = name;
    row[COL.EMAIL - 1] = email;
    row[COL.PHONE - 1] = phone;
    row[COL.AGE_RANGE - 1] = ageRange;
    row[COL.SLEEP_QUALITY - 1] = sleepQuality;
    row[COL.ENERGY_LEVELS - 1] = energyLevels;
    row[COL.MENTAL_CLARITY - 1] = mentalClarity;
    row[COL.STRESS_TOLERANCE - 1] = stressTolerance;
    row[COL.PHYSICAL_ACTIVITY - 1] = activityLevel;
    row[COL.BODY_CHANGES - 1] = bodyChanges;
    row[COL.LIBIDO - 1] = libido;
    row[COL.TEMP_SENSITIVITY_F - 1] = tempSensitivityF;
    row[COL.CYCLE_STATUS_F - 1] = cycleStatusF;
    row[COL.PMS_MOOD_F - 1] = pmsMoodF;
    row[COL.BLOATING_F - 1] = bloatingF;
    row[COL.STRENGTH_RECOVERY_M - 1] = strengthRecoveryM;
    row[COL.MORNING_DRIVE_M - 1] = morningDriveM;
    row[COL.MUSCLE_MASS_M - 1] = muscleMassM;
    row[COL.SEXUAL_PERFORMANCE_M - 1] = sexualPerformanceM;
    row[COL.BIO_BALANCE_SCORE - 1] = scorePercent;
    row[COL.LEAD_SCORE_SALES - 1] = leadScoreSales;
    row[COL.PROMO_CODE - 1] = promoCode;
    row[COL.PROMO_DESCRIPTION - 1] = promoDescription;
    row[COL.PROMO_REDEEMED - 1] = payload.promoRedeemed || "";
    row[COL.B12_OFFERED - 1] = payload.b12Offered || "";
    row[COL.B12_REDEEMED - 1] = payload.b12Redeemed || "";
    row[COL.LEAD_STATUS - 1] = payload.leadStatus || "New";
    row[COL.ASSIGNED_STAFF - 1] = payload.assignedStaff || "";
    row[COL.FOLLOW_UP_PRIORITY - 1] = payload.followUpPriority || "Medium";
    row[COL.LAST_OUTREACH_DATE - 1] = payload.lastOutreachDate || "";
    row[COL.NEXT_FOLLOW_UP_DATE - 1] = payload.nextFollowUpDate || "";
    row[COL.CONTACT_ATTEMPTS - 1] = payload.contactAttempts || 0;
    row[COL.LAST_ATTEMPT_METHOD - 1] = payload.lastAttemptMethod || "";
    row[COL.LAST_ATTEMPT_RESULT - 1] = payload.lastAttemptResult || "";
    row[COL.OPTED_OUT - 1] = payload.optedOut || "";
    row[COL.NOTES - 1] = payload.notes || "";
    row[COL.BOOKED_CONSULTATION - 1] = payload.bookedConsultation || "";
    row[COL.CONSULTATION_DATE - 1] = payload.consultationDate || "";
    row[COL.JOINED_MEMBERSHIP - 1] = payload.joinedMembership || "";
    row[COL.MEMBERSHIP_TYPE - 1] = payload.membershipType || "";
    row[COL.MEMBERSHIP_START_DATE - 1] = payload.membershipStartDate || "";
    row[COL.RETARGET_BUCKET - 1] = payload.coldLeadRetargetingBucket || payload.retargetingBucket || "";
    row[COL.TAGS - 1] = payload.tags || "";
    row[COL.RAW_ANSWERS - 1] = rawAnswers;
    row[COL.LEAD_TYPE - 1] = leadType;
    row[COL.FOLLOW_UP_SENT - 1] = payload.followUpSent || "";
    row[COL.FOLLOW_UP_72HR_TIMESTAMP - 1] = payload.followUp72hrTimestamp || "";
    row[COL.CLOSED_FROM - 1] = payload.closedFrom || "";

    ensureColumnExists(sheet, COL.FOLLOW_UP_72HR_TIMESTAMP);
    ensureColumnExists(sheet, COL.CLOSED_FROM);
    const insertedRowIndex = HEADER_ROW_INDEX + 1;
    sheet.insertRowsAfter(HEADER_ROW_INDEX, 1);
    const cleanedRow = row.map(cleanValue);
    sheet.getRange(insertedRowIndex, 1, 1, cleanedRow.length).setValues([cleanedRow]);
    sheet.setRowHeight(insertedRowIndex, 90);
    updateStatus(leadId, 2, "Running symptom matrix analysis");

    /*************************************************************
     *  GENERATE AI REPORT USING GEMINI 2.5 PRO
     *************************************************************/
    const prompt = buildGeminiPrompt({
      name,
      gender,
      ageRange,
      scorePercent,
      promoCode,
      source,
      sleepQuality,
      energyLevels,
      mentalClarity,
      stressTolerance,
      activityLevel,
      bodyChanges,
      libido,
      tempSensitivityF,
      cycleStatusF,
      pmsMoodF,
      bloatingF,
      strengthRecoveryM,
      morningDriveM,
      muscleMassM,
      sexualPerformanceM
    });
    updateStatus(leadId, 3, "Mapping your hormone pattern profile");

    let reportText;
    try {
      reportText = callGemini(prompt);
    } catch (geminiErr) {
      console.error("Gemini generation error", geminiErr);
      reportText = "<p>We were unable to auto-generate your expanded Bio-Balance report right now. A provider will review your assessment and follow up shortly.</p>";
    }
    updateStatus(leadId, 4, "Reviewing potential imbalance indicators");

    /*************************************************************
     *  BUILD HTML WRAPPER
     *************************************************************/
    const onscreenReportText = stripOnscreenPromotions(reportText);

    const reportHtml = buildReportHtml({
      name,
      gender,
      scorePercent,
      promoCode,
      reportText: onscreenReportText
    });
    updateStatus(leadId, 5, "Building your Bio-Balance Score");

    /*************************************************************
     *  EMAIL REPORT IN BODY (NO PDF ATTACHMENT)
     *************************************************************/
    if (email) {
      const emailHtml = composeReportEmailHtml({
        name: name,
        gender: gender,
        ageRange: ageRange,
        promoCode: promoCode,
        reportText: reportText
      });
      MailApp.sendEmail({
        to: email,
        subject: "Your Personalized Bio-Balance Report – White Tank Wellness",
        htmlBody: emailHtml
      });
    }

    updateStatus(leadId, 6, "Finalizing your clinical report");

    try {
      updateLeadTags({
        rowIndex: insertedRowIndex,
        sheetName: SUBMISSIONS_SHEET_NAME,
        manualTags: parseTagsInput(payload.tags || payload.manualTags || [])
      });
    } catch (tagError) {
      console.warn("Auto tag assignment failed for submission:", tagError);
    }

    /*************************************************************
     *  RETURN TO QUIZ FRONTEND
     *************************************************************/
    return sendCorsResponse(JSON.stringify({
      status: "success",
      reportHtml: reportHtml,
      leadId: leadId,
      leadType: leadType,
      name: name,
      email: email,
      scorePercent: scorePercent,
      promoCode: promoCode,
      rowIndex: insertedRowIndex
    }));

  } catch (err) {
    if (leadId) {
      updateStatus(leadId, 6, "Report generation failed");
    }
    return sendCorsResponse(JSON.stringify({
      status: "error",
      message: err.toString()
    }));
  }
}

/*************************************************************
 *  GEMINI PROMPT
 *************************************************************/
function buildGeminiPrompt(data) {
  return `
You are a functional medicine clinician at White Tank Wellness.

Generate ONE fully unified, premium, medical-grade report in clean HTML.  
Use ONLY <div>, <h2>, <h3>, <p>, <ul>, <li>, <hr>, and inline CSS.  
NO <html> or <body> tags.

The style must be:
- clean
- clinical
- beautifully designed
- uplifting & reassuring
- on-brand (colors: #0a466b, #99bed5, #e18d5a)
- premium and trustworthy
- mobile-friendly

-------------------------------------------
# OVERALL LAYOUT
-------------------------------------------
Wrap the entire report in:

<div style="font-family:Arial, sans-serif; max-width:800px; margin:0 auto; padding:24px; color:#222; line-height:1.6;">

-------------------------------------------
# LETTERHEAD (TOP OF REPORT)
-------------------------------------------
Insert this letterhead block as the first element inside the main wrapper:

<div style="
  display:flex;
  align-items:center;
  margin-bottom:24px;
  padding-bottom:16px;
  border-bottom:2px solid #0a466b20;
">

  <img src="https://cdn.prod.website-files.com/687e8476a1c54d5eea268c04/687e8476a1c54d5eea268cd0_sm-horz-logo.webp"
       alt="White Tank Wellness Logo"
       style="height:60px; width:auto; margin-right:16px;">

  <div style="font-size:14px; color:#0a466b; font-weight:600;">
    White Tank Wellness<br>
    208 N 4th St, Buckeye, AZ<br>
    602-761-9355<br>
    WhiteTankWellness.com
  </div>

</div>

-------------------------------------------
# WATERMARK SEAL (FAINT)
-------------------------------------------
A subtle, centered watermark reading "Bio-Balance Method™":

<div style="position:relative;">
  <div style="
    position:absolute;
    top:50%;
    left:50%;
    transform:translate(-50%, -50%);
    font-size:70px;
    color:#0a466b;
    opacity:0.06;
    white-space:nowrap;
    z-index:0;
  ">Bio-Balance Method™</div>

-------------------------------------------
# SECTION 1 — TITLE + SCORE BADGE
-------------------------------------------
At the top, create:

1. Large title:
<h2 style="color:#0a466b; margin-bottom:6px;">Bio-Balance Method™ Wellness Report</h2>

2. A gradient separator bar:
<div style="height:4px; background:linear-gradient(90deg,#0a466b,#99bed5,#e18d5a); margin:16px 0;"></div>

3. A score badge:
<div style="
  display:inline-block;
  padding:12px 20px;
  background:#e18d5a;
  color:white;
  border-radius:50px;
  font-weight:700;
  font-size:16px;
">Bio-Balance Score: ${data.scorePercent}/100</div>

4. A mini progress bar:
<div style="margin-top:10px; width:80%; background:#e5e8eb; height:10px; border-radius:6px;">
  <div style="
    width:${data.scorePercent}%;
    height:100%;
    background:#0a466b;
    border-radius:6px;
  "></div>
</div>
# PROMO CODE RULE (ONLY IF PROMO PROVIDED)
-------------------------------------------
${data.promoCode
  ? `
The promo code "${data.promoCode}" gives:
- **${autoPromoDescription(data.promoCode)}**

Always phrase it exactly as shown above.
`
  : `
<!-- No promotion selected; omit promo messaging -->
`
}

-------------------------------------------
# SECTION 2 — PATIENT SUMMARY BOX
-------------------------------------------
A clean clinical info box:

<div style="
  background:#f9fbfc;
  border:1px solid #d9e4ea;
  padding:18px;
  border-radius:10px;
  margin-top:24px;
  position:relative;
  z-index:2;
">
  <h3 style="margin-top:0; color:#0a466b;">Patient Summary</h3>
  <p><strong>Name:</strong> ${data.name || "Not provided"}</p>
  <p><strong>Gender:</strong> ${data.gender}</p>
  <p><strong>Age Range:</strong> ${data.ageRange || "Not specified"}</p>
  <p><strong>Promo Code:</strong> ${data.promoCode || "WTW25"}</p>
${(Array.isArray(data.symptoms) && data.symptoms.length > 0) 
  ? `  <p><strong>Main Concerns:</strong> ${data.symptoms.join(', ')}</p>` 
  : ''}
${data.notes 
  ? `  <p><strong>Provider Notes:</strong> ${data.notes}</p>` 
  : ''}
</div>

-------------------------------------------
# SECTION 3 — OVERVIEW OF THE SCORE
-------------------------------------------
<h3 style="color:#0a466b;">1. Understanding Your Bio-Balance Score</h3>
Write 2–3 paragraphs explaining what their score means — WITHOUT diagnosing.

-------------------------------------------
# SECTION 4 — HOW SYMPTOMS FIT TOGETHER
-------------------------------------------
<h3 style="color:#0a466b;">2. How Your Symptoms Fit Together</h3>
${(Array.isArray(data.symptoms) && data.symptoms.length > 0) 
  ? `**PRIMARY CONCERNS FROM INTAKE:** ${data.symptoms.join(', ')}
These are the patient's self-reported main concerns. Prioritize addressing these in your analysis.
` 
  : ''}
${data.notes 
  ? `**PROVIDER NOTES:** ${data.notes}
Consider these clinical observations when crafting your analysis.
` 
  : ''}

Discuss their specific symptoms:

- Sleep: ${data.sleepQuality}
- Energy: ${data.energyLevels}
- Mental clarity: ${data.mentalClarity}
- Stress tolerance: ${data.stressTolerance}
- Activity level: ${data.activityLevel}
- Body changes: ${data.bodyChanges}
- Libido: ${data.libido}

For female:
- Cycle status: ${data.cycleStatusF}
- PMS mood: ${data.pmsMoodF}
- Temperature sensitivity: ${data.tempSensitivityF}
- Bloating: ${data.bloatingF}

For male:
- Strength/recovery: ${data.strengthRecoveryM}
- Morning drive: ${data.morningDriveM}
- Muscle mass: ${data.muscleMassM}
- Sexual performance: ${data.sexualPerformanceM}

Write 2–4 paragraphs connecting their responses to general wellness patterns.

-------------------------------------------
# SECTION 5 — POSSIBLE WELLNESS PATTERNS
-------------------------------------------
<h3 style="color:#0a466b;">3. Possible Hormone & Stress Patterns</h3>
NEVER diagnose.  
Provide high-level phrasing like:  
- "This pattern may suggest shifts in…"  
- "These symptoms can sometimes reflect…"

-------------------------------------------
# SECTION 6 — PERSONALIZED RECOMMENDATIONS
-------------------------------------------
<h3 style="color:#0a466b;">4. Personalized Wellness Recommendations</h3>
Include:

<ul>
  <li>Sleep optimization</li>
  <li>Nervous-system calming practices</li>
  <li>Nutrition focusing on hormone support</li>
  <li>Movement and recovery</li>
  <li>Stress reduction strategies</li>
  <li>General lifestyle foundations</li>
</ul>

All neutral and non-medical.

-------------------------------------------
# SECTION 7 — NEXT STEPS WITH WTW
-------------------------------------------
<h3 style="color:#0a466b;">5. Your Next Steps</h3>
Include:

<p>You're welcome to schedule a consultation with a White Tank Wellness provider to review your results.</p>

Create a button:
<div style="margin:20px 0;">
  <a href="https://www.optimantra.com/optimus/patient/patientaccess/servicesall?pid=dkdSRlJnSnpWWWZmZ2J4Q3ExUUhCQT09&lid=RjBaSTBqc2tkV0FKSUVTRm9rS1k4UT09" target="_blank" style="
    background:#0a466b;
    color:white;
    padding:12px 24px;
    text-decoration:none;
    border-radius:6px;
    font-weight:600;
  ">Schedule Your Consultation</a>
</div>

Immediately after the button, add this premium QR code block:

<!-- PREMIUM QR CODE BLOCK -->
<div style="
  margin:40px auto;
  text-align:center;
  padding:24px;
  background:#f9fbfc;
  border:1px solid #d9e4ea;
  border-radius:14px;
  max-width:360px;
  box-shadow:0 4px 16px rgba(0,0,0,0.06);
">

  <!-- Ribbon Header -->
  <div style="
    background:linear-gradient(90deg,#0a466b,#99bed5);
    color:white;
    padding:10px 0;
    border-radius:10px 10px 0 0;
    font-weight:700;
    font-size:15px;
    letter-spacing:0.5px;
  ">
    Book Your Consultation
  </div>

  <!-- QR IMAGE -->
  <div style="
    background:white;
    padding:18px;
    border-radius:0 0 10px 10px;
  ">
    <img 
      src="https://api.qrserver.com/v1/create-qr-code/?size=220x220&color=10-70-107&bgcolor=f9-fb-fc&data=https%3A%2F%2Fwww.optimantra.com%2Foptimus%2Fpatient%2Fpatientaccess%2Fservicesall%3Fpid%3DdkdSRlJnSnpWWWZmZ2J4Q3ExUUhCQT09%26lid%3DRjBaSTBqc2tkV0FKSUVTRm9rS1k4UT09"
      alt="Scan to Schedule Consultation"
      style="
        width:220px;
        height:220px;
        border-radius:12px;
        border:4px solid #0a466b22;
        box-shadow:0 4px 14px rgba(0,0,0,0.08);
      ">
  </div>

  <!-- Scan Me Label -->
  <div style="
    margin-top:16px;
    font-size:14px;
    font-weight:600;
    color:#0a466b;
    letter-spacing:0.4px;
  ">
    Scan Me
  </div>

  <!-- Instructions -->
  <p style="font-size:13px; color:#555; margin-top:6px; line-height:1.5;">
    Open your phone's camera and tap the link to schedule your consultation
    with White Tank Wellness.
  </p>

</div>

-------------------------------------------
# END
-------------------------------------------

Close the main wrapper div.
</div>

Fill all narrative sections with warm, inclusive, professional text.
Do NOT sound robotic. Keep it optimistic and supportive.
  `;
}

/*************************************************************
 *  CALL GEMINI 2.5 PRO (NEW MODEL)
 *************************************************************/
function callGemini(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("GEMINI_API_KEY not set.");

  const url = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-pro:generateContent?key=" + apiKey;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  try {
    return json.candidates[0].content.parts[0].text;
  } catch (err) {
    return "<p>We were unable to generate your full report at this time. Please contact White Tank Wellness.</p>";
  }
}

/*************************************************************
 *  Strip CTA + QR blocks from on-screen report delivery
 *************************************************************/
function stripOnscreenPromotions(reportText) {
  if (!reportText) return reportText;

  const markers = [
    "margin:20px 0;",
    "margin: 20px 0;",
    "margin:40px auto;",
    "margin: 40px auto;"
  ];

  let cleaned = reportText;
  markers.forEach(marker => {
    let next = stripBlockByMarker(cleaned, marker);
    while (next !== cleaned) {
      cleaned = next;
      next = stripBlockByMarker(cleaned, marker);
    }
  });

  return cleaned.replace(/<!-- PREMIUM QR CODE BLOCK -->/g, "");
}

/*************************************************************
 *  Remove a full <div> block that contains a style marker
 *************************************************************/
function stripBlockByMarker(source, marker) {
  if (!source || !marker) return source;

  const markerIndex = source.indexOf(marker);
  if (markerIndex === -1) return source;

  let start = source.lastIndexOf("<div", markerIndex);
  if (start === -1) return source;

  let depth = 0;
  let cursor = start;

  while (cursor < source.length) {
    const nextOpen = source.indexOf("<div", cursor);
    const nextClose = source.indexOf("</div>", cursor);

    const hasOpen = nextOpen !== -1;
    const hasClose = nextClose !== -1;

    if (hasOpen && (!hasClose || nextOpen <= nextClose)) {
      depth++;
      cursor = nextOpen + 4;
      continue;
    }

    if (hasClose) {
      depth--;
      cursor = nextClose + 6;
      if (depth <= 0) {
        return source.slice(0, start) + source.slice(cursor);
      }
      continue;
    }

    break;
  }

  return source;
}

/*************************************************************
 *  HTML WRAPPER FOR ON-PAGE DISPLAY
 *************************************************************/
function buildReportHtml(opts) {
  const body = opts.reportText.replace(/\n/g, "<br>");

  return `
    <div id="aiGeneratedReport" style="
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
      max-width: 800px;
      margin: 20px auto;
      background: linear-gradient(135deg, #ffffff 0%, #f8fafb 100%);
      border-radius: 16px;
      box-shadow: 0 8px 32px rgba(10, 70, 107, 0.12);
      overflow: hidden;
      border: 1px solid rgba(10, 70, 107, 0.1);">
      
      <!-- Header Section -->
      <div style="
        background: linear-gradient(135deg, #0a466b 0%, #1a5a7f 100%);
        padding: 32px 28px;
        text-align: center;
        position: relative;
        overflow: hidden;">
        <div style="
          position: absolute;
          top: -50%;
          right: -10%;
          width: 300px;
          height: 300px;
          background: rgba(255, 255, 255, 0.05);
          border-radius: 50%;"></div>
        <h2 style="
          color: #ffffff;
          margin: 0 0 8px 0;
          font-size: 28px;
          font-weight: 700;
          letter-spacing: -0.5px;
          position: relative;
          z-index: 1;">White Tank Wellness</h2>
        <p style="
          color: rgba(255, 255, 255, 0.9);
          margin: 0;
          font-size: 16px;
          font-weight: 500;
          position: relative;
          z-index: 1;">Bio-Balance Method™ Report</p>
      </div>

      <!-- Patient Info Card -->
      <div style="padding: 0 28px 28px;">
        <div style="
          background: linear-gradient(135deg, #f0f7fb 0%, #e6f2f8 100%);
          border-left: 4px solid #0a466b;
          border-radius: 12px;
          padding: 24px;
          margin-bottom: 24px;
          box-shadow: 0 2px 8px rgba(10, 70, 107, 0.08);">
          <h3 style="
            margin: 0 0 16px 0;
            font-size: 16px;
            font-weight: 700;
            color: #0a466b;
            text-transform: uppercase;
            letter-spacing: 0.5px;">Patient Information</h3>
          <div style="display: grid; gap: 10px; font-size: 15px; line-height: 1.6;">
            <div><strong style="color: #0a466b; min-width: 120px; display: inline-block;">Name:</strong> <span style="color: #333;">${sanitize(opts.name)}</span></div>
            <div><strong style="color: #0a466b; min-width: 120px; display: inline-block;">Gender:</strong> <span style="color: #333;">${sanitize(opts.gender)}</span></div>
            <div><strong style="color: #0a466b; min-width: 120px; display: inline-block;">Promo Code:</strong> <span style="color: #333; background: #fff3e0; padding: 4px 12px; border-radius: 6px; font-weight: 600; letter-spacing: 1px;">${sanitize(opts.promoCode)}</span></div>
          </div>
        </div>

        <!-- Report Content -->
        <div style="
          background: #ffffff;
          padding: 28px;
          border-radius: 12px;
          border: 1px solid #e8eef2;
          font-size: 15px;
          line-height: 1.8;
          color: #2c3e50;
          box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);">
          ${body}
        </div>

        
      </div>
    </div>
  `;
}

function sanitize(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function composeReportEmailHtml(details) {
  const safeName = sanitize(details.name || "");
  const greetingName = safeName || "there";
  const gender = sanitize(details.gender || "Not specified");
  const ageRange = sanitize(details.ageRange || "Not specified");
  const promoCode = sanitize(details.promoCode || "WTW25");
  const reportText = details.reportText || "";

  return `
<div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif; background: linear-gradient(135deg, #e8f1f5 0%, #f5f8fa 100%); padding: 40px 20px;">

  <table width="100%" cellpadding="0" cellspacing="0"
    style="max-width: 680px; margin: 0 auto; background: #ffffff; border-radius: 16px;
           box-shadow: 0 10px 40px rgba(10, 70, 107, 0.15); overflow: hidden; border: 1px solid rgba(10, 70, 107, 0.1);">

    <!-- HEADER WITH GRADIENT -->
    <tr>
      <td style="background: linear-gradient(135deg, #0a466b 0%, #1a5a7f 100%); padding: 40px 32px; text-align: center; position: relative;">
        <div style="position: absolute; top: -30px; right: -30px; width: 180px; height: 180px; background: rgba(255, 255, 255, 0.05); border-radius: 50%;"></div>
        <h1 style="margin: 0 0 8px 0; font-size: 32px; font-weight: 800; color: #ffffff; letter-spacing: -0.5px; position: relative; z-index: 1;">White Tank Wellness</h1>
        <p style="margin: 0; font-size: 16px; color: rgba(255, 255, 255, 0.9); font-weight: 500; position: relative; z-index: 1;">🌟 Your Personalized Bio-Balance Report</p>
      </td>
    </tr>



    <!-- PATIENT SUMMARY CARD -->
    <tr>
      <td style="padding: 32px;">
        <div style="background: linear-gradient(135deg, #f0f7fb 0%, #e6f2f8 100%); border-left: 5px solid #0a466b; border-radius: 12px; padding: 24px; box-shadow: 0 4px 12px rgba(10, 70, 107, 0.08);">
          <h3 style="margin: 0 0 16px 0; font-size: 18px; font-weight: 700; color: #0a466b; text-transform: uppercase; letter-spacing: 0.5px;">📋 Patient Information</h3>
          <table width="100%" cellpadding="8" cellspacing="0" style="font-size: 15px; line-height: 1.6;">
            <tr>
              <td style="font-weight: 600; color: #0a466b; width: 140px; padding: 8px 0;">Name:</td>
              <td style="color: #2c3e50; padding: 8px 0;">${safeName || "Not provided"}</td>
            </tr>
            <tr>
              <td style="font-weight: 600; color: #0a466b; padding: 8px 0;">Gender:</td>
              <td style="color: #2c3e50; padding: 8px 0;">${gender}</td>
            </tr>
            <tr>
              <td style="font-weight: 600; color: #0a466b; padding: 8px 0;">Age Range:</td>
              <td style="color: #2c3e50; padding: 8px 0;">${ageRange}</td>
            </tr>
            <tr>
              <td style="font-weight: 600; color: #0a466b; padding: 8px 0;">Promo Code:</td>
              <td style="padding: 8px 0;">
                <span style="background: linear-gradient(135deg, #fff3e0 0%, #ffe8cc 100%); color: #d67a47; padding: 8px 16px; border-radius: 8px; font-weight: 700; letter-spacing: 1.5px; display: inline-block; border: 2px dashed #e18d5a;">${promoCode}</span>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>

    <!-- DECORATIVE SEPARATOR -->
    <tr><td style="height: 4px; background: linear-gradient(90deg, #99bed5 0%, #0a466b 100%);"></td></tr>

    <!-- WELCOME MESSAGE -->
    <tr>
      <td style="padding: 32px; color: #2c3e50; font-size: 16px; line-height: 1.8;">
        <p style="margin: 0 0 16px 0;">Hi <strong style="color: #0a466b;">${greetingName}</strong>,</p>
        <p style="margin: 0 0 16px 0;">Thank you for completing the <strong style="color: #0a466b;">Bio-Balance Method™ Wellness Assessment</strong>. Your personalized report is ready and waiting below.</p>
        <p style="margin: 0;">This comprehensive analysis is based on your unique responses and provides insights into patterns related to energy, sleep, mood, metabolism, and hormonal wellness. We're here to support your journey to optimal health.</p>
      </td>
    </tr>

    <!-- REPORT CONTENT WITH ENHANCED STYLING -->
    <tr>
      <td style="padding: 0 32px 32px;">
        <div style="background: #ffffff; border: 2px solid #e8eef2; border-radius: 12px; padding: 28px; font-size: 15px; color: #2c3e50; line-height: 1.8; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);">
          ${reportText}
        </div>
      </td>
    </tr>

    <!-- CALL TO ACTION SECTION -->
    <tr>
      <td style="padding: 32px; text-align: center; background: linear-gradient(135deg, #f8fafb 0%, #f0f4f6 100%);">
        <h3 style="margin: 0 0 16px 0; font-size: 22px; font-weight: 700; color: #0a466b;">Ready to Take the Next Step?</h3>
        <p style="margin: 0 0 24px 0; font-size: 15px; color: #6b7280; line-height: 1.6;">Schedule a consultation with our expert wellness team to discuss your personalized results.</p>
        
        <a href="https://www.optimantra.com/optimus/patient/patientaccess/servicesall?pid=dkdSRlJnSnpWWWZmZ2J4Q3ExUUhCQT09&lid=RjBaSTBqc2tkV0FKSUVTRm9rS1k4UT09" target="_blank"
          style="display: inline-block; background: linear-gradient(135deg, #0a466b 0%, #1a5a7f 100%); color: #ffffff; padding: 16px 32px; text-decoration: none; border-radius: 10px; font-weight: 700; font-size: 16px; box-shadow: 0 6px 20px rgba(10, 70, 107, 0.3); transition: all 0.3s ease;">
          📅 Schedule Your Consultation
        </a>
        
        <div style="margin-top: 24px; padding-top: 24px; border-top: 1px solid #d0dae0;">
          <p style="margin: 0; font-size: 15px; color: #6b7280;">Prefer to call? <strong style="color: #0a466b;">📞 602-761-9355</strong></p>
        </div>
      </td>
    </tr>

    <!-- FOOTER SEPARATOR -->
    <tr><td style="height:3px; background:linear-gradient(90deg,#e18d5a,#99bed5);"></td></tr>

    <!-- FOOTER -->
    <tr>
      <td style="padding:20px; text-align:center; background:#f7f9fb; font-size:12px; color:#666; position:relative; z-index:2;">

        <!-- Social Links -->
        <div style="margin-bottom:12px;">
          <a href="https://www.facebook.com/whitetankwellness" target="_blank"
            style="margin-right:10px; color:#0a466b; font-weight:600; text-decoration:none;">Facebook</a>

          <a href="https://www.instagram.com/whitetankwellness" target="_blank"
            style="margin-left:10px; color:#0a466b; font-weight:600; text-decoration:none;">Instagram</a>
        </div>

        White Tank Wellness<br>
        208 N 4th St • Buckeye, AZ<br><br>

        <span style="color:#999;">This report is educational only and is not medical advice.</span>
      </td>
    </tr>

  </table>
</div>
`;
}

/*************************************************************
 *  PATIENT LEAD SIDEBAR + CALL LOGGING SUPPORT
 *************************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("White Tank Wellness")
    .addItem("Open Patient Lead Card", "openLeadSidebar")
    .addItem("Open Internal Intake", "openInternalIntakeSidebar")
    .addToUi();
}

function openLeadSidebar() {
  const html = HtmlService.createTemplateFromFile("leadSidebar")
    .evaluate()
    .setTitle("Patient Lead Card");

  SpreadsheetApp.getUi().showSidebar(html);
}

function openInternalIntakeSidebar(prefillLeadFromClient) {
  let prefillLead = null;

  if (prefillLeadFromClient && typeof prefillLeadFromClient === "object") {
    prefillLead = prefillLeadFromClient;
  }

  if (!prefillLead) {
    const cache = CacheService.getUserCache();
    const cached = cache.get("currentLead");
    if (cached) {
      try {
        prefillLead = JSON.parse(cached);
      } catch (err) {
        prefillLead = null;
      }
    }
  }

  if (!prefillLead) {
    try {
      prefillLead = getSelectedLead();
    } catch (selectionErr) {
      prefillLead = null;
    }
  }

  if (prefillLead) {
    injectLeadIntoSidebar(prefillLead);
  }

  const html = HtmlService.createTemplateFromFile("InternalIntakeSidebar");
  html.prefillLead = prefillLead;

  const evaluated = html
    .evaluate()
    .setTitle("Internal Intake Form");

  SpreadsheetApp.getUi().showSidebar(evaluated);
}

function getSelectedLead() {
  const activeRange = SpreadsheetApp.getActiveRange();
  if (!activeRange) {
    throw new Error("Select a row inside the 'Customers' sheet first.");
  }

  const selectedSheet = activeRange.getSheet();
  const selectedSheetName = selectedSheet ? selectedSheet.getName() : "";
  if (!selectedSheet || ![SUBMISSIONS_SHEET_NAME, CUSTOMERS_SHEET_NAME].includes(selectedSheetName)) {
    throw new Error("Please highlight a row in the 'Customers' sheet.");
  }

  const rowIndex = activeRange.getRow();
  if (rowIndex <= HEADER_ROW_INDEX) {
    throw new Error("Select a data row below the header row.");
  }

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(selectedSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + selectedSheetName + "' not found.");
  }

  const lastColumn = sheet.getLastColumn() || COL.CLOSED_FROM;
  const values = sheet.getRange(rowIndex, 1, 1, lastColumn).getValues()[0];

  const lead = buildLeadObjectFromRow(rowIndex, values);
  lead.sheetName = selectedSheetName;
  return lead;
}

function updateLeadCallLog(payload) {
  if (!payload) {
    throw new Error("A valid payload is required.");
  }

  // Get row from payload or active selection
  let rowIndex;
  if (payload.rowIndex) {
    rowIndex = Number(payload.rowIndex);
  } else {
    const activeRange = SpreadsheetApp.getActiveRange();
    if (!activeRange) {
      throw new Error("No row selected.");
    }
    rowIndex = activeRange.getRow();
  }

  const method = (payload.method || "Contact").trim();
  const note = (payload.note || "").trim();

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const targetSheetName = payload.sheetName || CUSTOMERS_SHEET_NAME;
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  ensureColumnExists(sheet, COL.NEXT_FOLLOW_UP_DATE);

  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd/yy HH:mm");
  const userName = Session.getActiveUser().getEmail().split("@")[0] || "Staff";

  const followUpDetails = payload.followUpDetails || null;
  const markFollowUpComplete = Boolean(payload.markFollowUpComplete);
  const sanitizedFollowUp = followUpDetails ? {
    note: followUpDetails.note || "",
    dueInput: followUpDetails.dueInput || "",
    action: followUpDetails.action || "",
    dueIso: followUpDetails.dueIso || ""
  } : null;

  // Update Last Outreach Date
  sheet.getRange(rowIndex, COL.LAST_OUTREACH_DATE).setValue(now);

  // Update Last Attempt Method
  sheet.getRange(rowIndex, COL.LAST_ATTEMPT_METHOD).setValue(method);

  // Update Last Attempt Result
  sheet.getRange(rowIndex, COL.LAST_ATTEMPT_RESULT).setValue(note);

  // Increment Contact Attempts
  const currentAttempts = Number(sheet.getRange(rowIndex, COL.CONTACT_ATTEMPTS).getValue() || 0);
  sheet.getRange(rowIndex, COL.CONTACT_ATTEMPTS).setValue(currentAttempts + 1);

  // Append to Notes
  const existingNotes = sheet.getRange(rowIndex, COL.NOTES).getValue() || "";
  const formattedNote = "[" + timestamp + " - " + userName + "] " + method + ": " + note;
  const updatedNotes = existingNotes ? existingNotes + "\n" + formattedNote : formattedNote;
  sheet.getRange(rowIndex, COL.NOTES).setValue(updatedNotes);

  if (sanitizedFollowUp && (sanitizedFollowUp.note || sanitizedFollowUp.dueInput)) {
    let followUpDateValue = sanitizedFollowUp.dueInput;
    if (sanitizedFollowUp.dueIso) {
      const isoDate = new Date(sanitizedFollowUp.dueIso);
      if (!isNaN(isoDate.getTime())) {
        followUpDateValue = isoDate;
      }
    }
    if (followUpDateValue) {
      sheet.getRange(rowIndex, COL.NEXT_FOLLOW_UP_DATE).setValue(followUpDateValue);
    }
    setFollowUpHighlight(sheet, rowIndex, true);
  }

  if (markFollowUpComplete) {
    sheet.getRange(rowIndex, COL.NEXT_FOLLOW_UP_DATE).setValue("");
    setFollowUpHighlight(sheet, rowIndex, false);
  }

  const lead = getSelectedLeadFromRow(rowIndex, targetSheetName);
  injectLeadIntoSidebar(lead);

  return {
    success: true,
    method: method,
    note: note
  };
}

function setFollowUpHighlight(sheet, rowIndex, isOpen) {
  if (!sheet || !rowIndex) {
    return;
  }
  const lastCol = sheet.getLastColumn();
  if (!lastCol) {
    return;
  }
  const range = sheet.getRange(rowIndex, 1, 1, lastCol);
  range.setBackground(isOpen ? "#fff9c4" : "#ffffff");
}

function formatSubmittedDate(value) {
  if (!(value instanceof Date)) {
    return "";
  }
  return Utilities.formatDate(value, Session.getScriptTimeZone(), "MMM d, yyyy h:mm a");
}

function isTruthy(value) {
  if (value === true) return true;
  if (typeof value === "number") return value === 1;
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    return normalized === "true" || normalized === "yes" || normalized === "1";
  }
  return false;
}

// ----- NEW SIDEBAR REAL-TIME UPDATE SYSTEM ----- //

/*************************************************************
 *  onSelectionChange TRIGGER - Auto-update sidebar on row change
 *************************************************************/
function onSelectionChange(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (!sheet) return;

    const sheetName = sheet.getName();
    if (![SUBMISSIONS_SHEET_NAME, CUSTOMERS_SHEET_NAME].includes(sheetName)) return;

    const rowIndex = e.range.getRow();
    if (rowIndex <= HEADER_ROW_INDEX) return;

    // Get lead data and inject into sidebar
    const lead = getSelectedLeadFromRow(rowIndex, sheetName);
    injectLeadIntoSidebar(lead);
  } catch (err) {
    console.error("onSelectionChange error:", err);
  }
}

/*************************************************************
 *  getCurrentRowData - Returns active row data from Customers sheet
 *************************************************************/
function getCurrentRowData() {
  try {
    const activeRange = SpreadsheetApp.getActiveRange();
    if (!activeRange) {
      throw new Error("Select a row inside the '" + CUSTOMERS_SHEET_NAME + "' sheet first.");
    }

    const sheet = activeRange.getSheet();
    if (!sheet || sheet.getName() !== CUSTOMERS_SHEET_NAME) {
      throw new Error("Please highlight a row in the '" + CUSTOMERS_SHEET_NAME + "' sheet.");
    }

    const rowIndex = activeRange.getRow();
    if (rowIndex <= HEADER_ROW_INDEX) {
      throw new Error("Select a data row below the header row.");
    }

    const lead = getSelectedLeadFromRow(rowIndex, CUSTOMERS_SHEET_NAME);
    injectLeadIntoSidebar(lead);
    return lead;
  } catch (err) {
    console.error("getCurrentRowData error:", err);
    throw err;
  }
}

/*************************************************************
 *  getSelectedLeadFromRow - Returns lead object from row number
 *************************************************************/
function getSelectedLeadFromRow(rowIndex, sheetName) {
  if (!rowIndex || rowIndex <= HEADER_ROW_INDEX) {
    throw new Error("Invalid row index provided.");
  }

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const targetSheetName = sheetName || SUBMISSIONS_SHEET_NAME;
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  const lastColumn = sheet.getLastColumn() || COL.CLOSED_FROM;
  const values = sheet.getRange(rowIndex, 1, 1, lastColumn).getValues()[0];

  const lead = buildLeadObjectFromRow(rowIndex, values);
  lead.sheetName = targetSheetName;
  lead.sheetLastRow = sheet.getLastRow();
  lead.firstDataRow = HEADER_ROW_INDEX + 1;
  return lead;
}

/*************************************************************
 *  buildLeadObjectFromRow - Shared helper to construct lead object
 *  Updated to match new header structure
 *************************************************************/
function buildLeadObjectFromRow(rowIndex, values) {
  const membershipStartCell = values[COL.MEMBERSHIP_START_DATE - 1];
  const rawTagsCell = values[COL.TAGS - 1] || "";
  const parsedStoredTags = parseTagsInput(rawTagsCell);

  const lead = {
    rowIndex: rowIndex,
    // Core fields
    submittedAt: formatSubmittedDate(values[COL.TIMESTAMP - 1]),
    leadId: values[COL.LEAD_ID - 1] || generateLeadId(rowIndex),
    source: values[COL.SOURCE - 1] || "",
    // UTM fields
    utmCampaign: values[COL.UTM_CAMPAIGN - 1] || "",
    utmAdSet: values[COL.UTM_AD_SET - 1] || "",
    utmKeyword: values[COL.UTM_KEYWORD - 1] || "",
    utmLandingPage: values[COL.UTM_LANDING_PAGE - 1] || "",
    // Contact info
    gender: values[COL.GENDER - 1] || "",
    name: values[COL.NAME - 1] || "",
    email: values[COL.EMAIL - 1] || "",
    phone: values[COL.PHONE - 1] || "",
    ageRange: values[COL.AGE_RANGE - 1] || "",
    // Symptoms - shared
    sleepQuality: values[COL.SLEEP_QUALITY - 1] || "",
    energyLevels: values[COL.ENERGY_LEVELS - 1] || "",
    mentalClarity: values[COL.MENTAL_CLARITY - 1] || "",
    stressTolerance: values[COL.STRESS_TOLERANCE - 1] || "",
    activityLevel: values[COL.PHYSICAL_ACTIVITY - 1] || "",
    bodyChanges: values[COL.BODY_CHANGES - 1] || "",
    libido: values[COL.LIBIDO - 1] || "",
    // Symptoms - female
    tempSensitivityF: values[COL.TEMP_SENSITIVITY_F - 1] || "",
    cycleStatusF: values[COL.CYCLE_STATUS_F - 1] || "",
    pmsMoodF: values[COL.PMS_MOOD_F - 1] || "",
    bloatingF: values[COL.BLOATING_F - 1] || "",
    // Symptoms - male
    strengthRecoveryM: values[COL.STRENGTH_RECOVERY_M - 1] || "",
    morningDriveM: values[COL.MORNING_DRIVE_M - 1] || "",
    muscleMassM: values[COL.MUSCLE_MASS_M - 1] || "",
    sexualPerformanceM: values[COL.SEXUAL_PERFORMANCE_M - 1] || "",
    // Scores
    scorePercent: Number(values[COL.BIO_BALANCE_SCORE - 1] || 0),
    salesLeadScore: Number(values[COL.LEAD_SCORE_SALES - 1] || 0),
    // Promo
    promoCode: values[COL.PROMO_CODE - 1] || "",
    promoDescription: values[COL.PROMO_DESCRIPTION - 1] || autoPromoDescription(values[COL.PROMO_CODE - 1] || ""),
    promoRedeemed: isTruthy(values[COL.PROMO_REDEEMED - 1]),
    // B12 tracking
    b12Offered: isTruthy(values[COL.B12_OFFERED - 1]),
    b12Redeemed: isTruthy(values[COL.B12_REDEEMED - 1]),
    // Lead management
    leadStatus: values[COL.LEAD_STATUS - 1] || "New",
    assignedStaff: values[COL.ASSIGNED_STAFF - 1] || "",
    followUpPriority: values[COL.FOLLOW_UP_PRIORITY - 1] || "Medium",
    // Outreach tracking
    lastOutreachDate: formatSubmittedDate(values[COL.LAST_OUTREACH_DATE - 1]),
    nextFollowUpDate: formatSubmittedDate(values[COL.NEXT_FOLLOW_UP_DATE - 1]),
    contactAttempts: Number(values[COL.CONTACT_ATTEMPTS - 1] || 0),
    lastAttemptMethod: values[COL.LAST_ATTEMPT_METHOD - 1] || "",
    lastAttemptResult: values[COL.LAST_ATTEMPT_RESULT - 1] || "",
    optedOut: isTruthy(values[COL.OPTED_OUT - 1]),
    notes: values[COL.NOTES - 1] || "",
    // Consultation
    bookedConsultation: isTruthy(values[COL.BOOKED_CONSULTATION - 1]),
    consultationDate: formatSubmittedDate(values[COL.CONSULTATION_DATE - 1]),
    consultationDateIso: values[COL.CONSULTATION_DATE - 1] instanceof Date ? values[COL.CONSULTATION_DATE - 1].toISOString() : "",
    consultationDateRaw: values[COL.CONSULTATION_DATE - 1] || "",
    // Membership
    joinedMembership: isTruthy(values[COL.JOINED_MEMBERSHIP - 1]),
    membershipType: values[COL.MEMBERSHIP_TYPE - 1] || "",
    membershipStartDate: formatSubmittedDate(membershipStartCell),
    membershipStartDateIso: membershipStartCell instanceof Date ? membershipStartCell.toISOString() : "",
    membershipStartDateRaw: membershipStartCell,
    // Marketing
    retargetBucket: values[COL.RETARGET_BUCKET - 1] || "",
    tagsCellValue: rawTagsCell,
    rawAnswers: values[COL.RAW_ANSWERS - 1] || "",
    leadType: values[COL.LEAD_TYPE - 1] || "Quiz Lead",
    followUpSent: isTruthy(values[COL.FOLLOW_UP_SENT - 1]),
    followUp72hrTimestamp: formatSubmittedDate(values[COL.FOLLOW_UP_72HR_TIMESTAMP - 1]),
    closedFrom: values[COL.CLOSED_FROM - 1] || ""
  };

  const autoTags = computeAutoTagsFromLead(lead);
  const manualTags = deriveManualTags(parsedStoredTags, autoTags);
  const mergedTags = mergeAutoAndManualTags(autoTags, manualTags);

  lead.manualTags = manualTags;
  lead.autoTags = autoTags;
  lead.tags = mergedTags;
  lead.allTags = mergedTags.slice();
  lead.tagList = mergedTags.slice();
  lead.tagsString = mergedTags.join(", ");

  // Parse symptoms array from rawAnswers if available (for internal intake)
  if (lead.rawAnswers && typeof lead.rawAnswers === 'string') {
    try {
      const parsed = JSON.parse(lead.rawAnswers);
      if (Array.isArray(parsed.symptoms)) {
        lead.symptoms = parsed.symptoms;
      }
    } catch (e) {
      // If parsing fails, leave symptoms undefined
    }
  }

  return lead;
}

/*************************************************************
 *  generateLeadId - Create unique lead ID if not present
 *************************************************************/
function generateLeadId(rowIndex) {
  const timestamp = new Date().getTime().toString(36).toUpperCase();
  return "WTW-" + timestamp + "-" + rowIndex;
}

/*************************************************************
 *  injectLeadIntoSidebar - Calls client-side updateSidebarWithLead
 *************************************************************/
function injectLeadIntoSidebar(lead) {
  try {
    // HtmlService in sandboxed mode requires google.script.run from client
    // This function is called server-side and will trigger a client callback
    // The sidebar HTML must have updateSidebarWithLead(lead) defined
    
    // For real-time updates, we use CacheService to store the lead
    // and the sidebar polls for changes
    const cache = CacheService.getUserCache();
    cache.put("currentLead", JSON.stringify(lead), 300); // 5 minute cache
    
    return lead;
  } catch (err) {
    console.error("injectLeadIntoSidebar error:", err);
    return null;
  }
}

/*************************************************************
 *  getCurrentLeadFromCache - Sidebar polling endpoint
 *************************************************************/
function getCurrentLeadFromCache() {
  try {
    const cache = CacheService.getUserCache();
    const cachedLead = cache.get("currentLead");
    if (cachedLead) {
      return JSON.parse(cachedLead);
    }
    return null;
  } catch (err) {
    console.error("getCurrentLeadFromCache error:", err);
    return null;
  }
}

/*************************************************************
 *  updateLeadStatus - Update lead status in spreadsheet
 *************************************************************/
function updateLeadStatus(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);
  const leadStatus = (payload.leadStatus || "").trim();

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const targetSheetName = payload.sheetName || CUSTOMERS_SHEET_NAME;
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  // Ensure we have enough columns
  ensureColumnExists(sheet, COL.LEAD_STATUS);

  sheet.getRange(rowIndex, COL.LEAD_STATUS).setValue(leadStatus);

  // Update cache with new data
  const lead = getSelectedLeadFromRow(rowIndex, targetSheetName);
  injectLeadIntoSidebar(lead);

  try {
    updateLeadTags({
      rowIndex: rowIndex,
      sheetName: targetSheetName,
      manualTags: lead.manualTags || []
    });
  } catch (tagError) {
    console.warn("Auto tag refresh failed after status update:", tagError);
  }

  return {
    success: true,
    rowIndex: rowIndex,
    leadStatus: leadStatus
  };
}

/*************************************************************
 *  updateFollowUpPriority - Update follow-up priority
 *************************************************************/
function updateFollowUpPriority(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);
  const followUpPriority = (payload.followUpPriority || "").trim();

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const targetSheetName = payload.sheetName || CUSTOMERS_SHEET_NAME;
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  ensureColumnExists(sheet, COL.FOLLOW_UP_PRIORITY);

  sheet.getRange(rowIndex, COL.FOLLOW_UP_PRIORITY).setValue(followUpPriority);

  // Update cache
  const lead = getSelectedLeadFromRow(rowIndex, targetSheetName);
  injectLeadIntoSidebar(lead);

  return {
    success: true,
    rowIndex: rowIndex,
    followUpPriority: followUpPriority
  };
}

/*************************************************************
 *  appendLeadNotes - Append notes with timestamp
 *************************************************************/
function appendLeadNotes(notesPayload) {
  let rowIndex, newNote;
  let targetSheetName = CUSTOMERS_SHEET_NAME;

  // Handle both object payload and string note
  if (typeof notesPayload === "string") {
    // Called with just a string - get row from active selection
    const activeRange = SpreadsheetApp.getActiveRange();
    if (!activeRange) {
      throw new Error("No row selected.");
    }
    rowIndex = activeRange.getRow();
    newNote = notesPayload;
  } else if (notesPayload && notesPayload.rowIndex) {
    rowIndex = Number(notesPayload.rowIndex);
    newNote = (notesPayload.note || notesPayload.notes || "").trim();
    if (notesPayload.sheetName) {
      targetSheetName = notesPayload.sheetName;
    }
  } else {
    throw new Error("Invalid notes payload.");
  }

  if (!newNote) {
    throw new Error("Note content is required.");
  }

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  ensureColumnExists(sheet, COL.NOTES);

  // Get existing notes and append
  const existingNotes = sheet.getRange(rowIndex, COL.NOTES).getValue() || "";
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yy HH:mm");
  const userName = Session.getActiveUser().getEmail().split("@")[0] || "Staff";
  
  const formattedNote = "[" + timestamp + " - " + userName + "] " + newNote;
  const updatedNotes = existingNotes ? existingNotes + "\n" + formattedNote : formattedNote;

  sheet.getRange(rowIndex, COL.NOTES).setValue(updatedNotes);

  // Also increment contact attempts if this is an outreach log
  if (notesPayload && notesPayload.isOutreach) {
    const currentAttempts = Number(sheet.getRange(rowIndex, COL.CONTACT_ATTEMPTS).getValue() || 0);
    sheet.getRange(rowIndex, COL.CONTACT_ATTEMPTS).setValue(currentAttempts + 1);
  }

  // Update cache
  const lead = getSelectedLeadFromRow(rowIndex, targetSheetName);
  injectLeadIntoSidebar(lead);

  return {
    success: true,
    rowIndex: rowIndex,
    notes: updatedNotes
  };
}

/*************************************************************
 *  calculateSalesLeadScore - Compute weighted lead score
 *************************************************************/
function calculateSalesLeadScore(scorePercent, symptomData) {
  let salesScore = 0;

  // Base score from quiz (0-100 scale, weight 40%)
  const quizScore = Number(scorePercent) || 0;
  salesScore += (quizScore / 100) * 40;

  // Symptom severity indicators (weight 30%)
  const severityKeywords = ["severe", "very", "extreme", "always", "constantly", "never"];
  const moderateKeywords = ["often", "frequently", "sometimes", "moderate"];
  
  if (symptomData) {
    const symptomString = JSON.stringify(symptomData).toLowerCase();
    
    // Check for high-severity indicators
    let severityCount = 0;
    severityKeywords.forEach(function(keyword) {
      if (symptomString.includes(keyword)) severityCount++;
    });
    salesScore += Math.min(severityCount * 5, 20);

    // Check for moderate indicators
    let moderateCount = 0;
    moderateKeywords.forEach(function(keyword) {
      if (symptomString.includes(keyword)) moderateCount++;
    });
    salesScore += Math.min(moderateCount * 2.5, 10);
  }

  // Age range factor (weight 15%)
  if (symptomData && symptomData.ageRange) {
    const ageRange = symptomData.ageRange.toLowerCase();
    if (ageRange.includes("40") || ageRange.includes("50") || ageRange.includes("60")) {
      salesScore += 15; // Prime demographic for hormone therapy
    } else if (ageRange.includes("30")) {
      salesScore += 10;
    } else {
      salesScore += 5;
    }
  }

  // Contact info completeness (weight 15%)
  if (symptomData) {
    if (symptomData.email) salesScore += 5;
    if (symptomData.phone) salesScore += 5;
    if (symptomData.name) salesScore += 5;
  }

  return Math.min(Math.round(salesScore), 100);
}

/*************************************************************
 *  autoPromoDescription - Return promo code description
 *************************************************************/
function autoPromoDescription(promoCode) {
  const code = (promoCode || "").toUpperCase().trim();
  
  const promoMap = {
    "WTW25": "$25 off your first month of membership",
    "WELCOME25": "$25 off your first month of membership",
    "BIOBALANCE": "$25 off your first month of membership",
    "FRIEND25": "$25 off your first month of membership (Referral)",
    "VIP50": "$50 off your first month of membership (VIP)",
    "HOLIDAY": "$25 off + Free consultation",
    "SPRING2024": "$25 off your first month of membership",
    "NEWYEAR": "$25 off your first month of membership"
  };

  return promoMap[code] || "$25 off your first month of membership";
}

/*************************************************************
 *  autoRetargetBucket - Categorize lead for retargeting
 *************************************************************/
function autoRetargetBucket(leadObj) {
  if (!leadObj) return "Unknown";

  const score = Number(leadObj.scorePercent) || 0;
  const status = (leadObj.leadStatus || "").toLowerCase();
  const attempts = Number(leadObj.contactAttempts) || 0;
  const daysSinceSubmit = calculateDaysSince(leadObj.submittedAt);

  // Already converted
  if (status === "converted" || status === "member" || status === "patient") {
    return "Converted";
  }

  // Lost leads
  if (status === "lost" || status === "unqualified" || status === "do not contact") {
    return "Excluded";
  }

  // High intent - recent, high score, low contact attempts
  if (score >= 70 && daysSinceSubmit <= 7 && attempts <= 2) {
    return "Hot Lead - Immediate Follow-up";
  }

  // Warm leads - decent score, contacted but not converted
  if (score >= 50 && daysSinceSubmit <= 30) {
    return "Warm Lead - Nurture Sequence";
  }

  // Re-engagement needed
  if (daysSinceSubmit > 30 && daysSinceSubmit <= 90) {
    return "Re-engagement Campaign";
  }

  // Cold leads
  if (daysSinceSubmit > 90) {
    return "Cold Lead - Long-term Nurture";
  }

  // Low score leads
  if (score < 40) {
    return "Low Priority - Educational Content";
  }

  return "Standard Follow-up";
}

/*************************************************************
 *  calculateDaysSince - Helper for date calculations
 *************************************************************/
function calculateDaysSince(dateString) {
  if (!dateString) return 999;
  
  try {
    const submitDate = new Date(dateString);
    const now = new Date();
    const diffTime = Math.abs(now - submitDate);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays;
  } catch (err) {
    return 999;
  }
}

/*************************************************************
 *  ensureColumnExists - Expand sheet if needed
 *************************************************************/
function ensureColumnExists(sheet, columnIndex) {
  const currentColumns = sheet.getMaxColumns();
  if (columnIndex > currentColumns) {
    sheet.insertColumnsAfter(currentColumns, columnIndex - currentColumns);
  }
}

/*************************************************************
 *  updateAssignedStaff - Assign staff member to lead
 *************************************************************/
function updateAssignedStaff(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);
  const assignedStaff = (payload.assignedStaff || "").trim();

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet 'Submissions' not found.");
  }

  ensureColumnExists(sheet, COL.ASSIGNED_STAFF);
  sheet.getRange(rowIndex, COL.ASSIGNED_STAFF).setValue(assignedStaff);

  const lead = getSelectedLeadFromRow(rowIndex);
  injectLeadIntoSidebar(lead);

  return {
    success: true,
    rowIndex: rowIndex,
    assignedStaff: assignedStaff
  };
}

/*************************************************************
 *  updateConsultationInfo - Update consultation details
 *************************************************************/
function updateConsultationInfo(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const targetSheetName = payload.sheetName || SUBMISSIONS_SHEET_NAME;
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  ensureColumnExists(sheet, COL.CONSULTATION_DATE);

  if (payload.bookedConsultation !== undefined) {
    sheet.getRange(rowIndex, COL.BOOKED_CONSULTATION).setValue(payload.bookedConsultation ? "TRUE" : "FALSE");
  }
  if (payload.consultationDate) {
    sheet.getRange(rowIndex, COL.CONSULTATION_DATE).setValue(new Date(payload.consultationDate));
  }

  const lead = getSelectedLeadFromRow(rowIndex, targetSheetName);
  injectLeadIntoSidebar(lead);

  try {
    updateLeadTags({
      rowIndex: rowIndex,
      sheetName: targetSheetName,
      manualTags: lead.manualTags || []
    });
  } catch (tagError) {
    console.warn("Auto tag refresh failed after membership update:", tagError);
  }

  return {
    success: true,
    rowIndex: rowIndex
  };
}

/*************************************************************
 *  updateMembershipInfo - Update membership details
 *************************************************************/
function updateMembershipInfo(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const targetSheetName = payload.sheetName || SUBMISSIONS_SHEET_NAME;
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  ensureColumnExists(sheet, COL.MEMBERSHIP_START_DATE);

  if (payload.joinedMembership !== undefined) {
    sheet.getRange(rowIndex, COL.JOINED_MEMBERSHIP).setValue(payload.joinedMembership ? "TRUE" : "FALSE");
  }
  if (payload.membershipType) {
    sheet.getRange(rowIndex, COL.MEMBERSHIP_TYPE).setValue(payload.membershipType);
  }
  if (payload.membershipStartDate) {
    sheet.getRange(rowIndex, COL.MEMBERSHIP_START_DATE).setValue(new Date(payload.membershipStartDate));
  }

  const lead = getSelectedLeadFromRow(rowIndex, targetSheetName);
  injectLeadIntoSidebar(lead);

  try {
    updateLeadTags({
      rowIndex: rowIndex,
      sheetName: targetSheetName,
      manualTags: lead.manualTags || []
    });
  } catch (tagError) {
    console.warn("Auto tag refresh failed after membership details update:", tagError);
  }

  return {
    success: true,
    rowIndex: rowIndex
  };
}

/*************************************************************
 *  updateTags - Update lead tags
 *************************************************************/
function updateLeadTags(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);
  const manualTags = parseTagsInput(
    payload.manualTags !== undefined ? payload.manualTags : payload.tags
  );
  const targetSheetName = payload.sheetName || SUBMISSIONS_SHEET_NAME;

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  ensureColumnExists(sheet, COL.TAGS);

  const currentLead = getSelectedLeadFromRow(rowIndex, targetSheetName);
  const autoTags = computeAutoTagsFromLead(currentLead);
  const mergedTags = mergeAutoAndManualTags(autoTags, manualTags);
  const tagString = mergedTags.join(", ");

  sheet.getRange(rowIndex, COL.TAGS).setValue(tagString);

  const updatedLead = getSelectedLeadFromRow(rowIndex, targetSheetName);
  injectLeadIntoSidebar(updatedLead);

  return {
    success: true,
    rowIndex: rowIndex,
    manualTags: updatedLead.manualTags || manualTags,
    autoTags: updatedLead.autoTags || autoTags,
    tags: updatedLead.tags || mergedTags,
    allTags: updatedLead.allTags || mergedTags,
    tagString: updatedLead.tagsString || tagString,
    lead: updatedLead
  };
}

function updateTags(payload) {
  return updateLeadTags(payload || {});
}

/*************************************************************
 *  updateNextFollowUpDate - Set next follow-up date
 *************************************************************/
function updateNextFollowUpDate(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet 'Submissions' not found.");
  }

  ensureColumnExists(sheet, COL.NEXT_FOLLOW_UP_DATE);

  if (payload.nextFollowUpDate) {
    sheet.getRange(rowIndex, COL.NEXT_FOLLOW_UP_DATE).setValue(new Date(payload.nextFollowUpDate));
  }

  const lead = getSelectedLeadFromRow(rowIndex);
  injectLeadIntoSidebar(lead);

  return {
    success: true,
    rowIndex: rowIndex
  };
}

/*************************************************************
 *  updateOptedOut - Mark lead as opted out
 *************************************************************/
function updateOptedOut(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);
  const optedOut = Boolean(payload.optedOut);

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet 'Submissions' not found.");
  }

  ensureColumnExists(sheet, COL.OPTED_OUT);
  sheet.getRange(rowIndex, COL.OPTED_OUT).setValue(optedOut ? "TRUE" : "FALSE");

  // Also update lead status if opted out
  if (optedOut) {
    sheet.getRange(rowIndex, COL.LEAD_STATUS).setValue("Do Not Contact");
  }

  const lead = getSelectedLeadFromRow(rowIndex);
  injectLeadIntoSidebar(lead);

  return {
    success: true,
    rowIndex: rowIndex,
    optedOut: optedOut
  };
}

/*************************************************************
 *  markPromoRedeemed - Mark promo code as redeemed
 *************************************************************/
function markPromoRedeemed(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const targetSheetName = payload.sheetName || SUBMISSIONS_SHEET_NAME;
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  ensureColumnExists(sheet, COL.PROMO_REDEEMED);
  sheet.getRange(rowIndex, COL.PROMO_REDEEMED).setValue("TRUE");

  const lead = getSelectedLeadFromRow(rowIndex, targetSheetName);
  injectLeadIntoSidebar(lead);

  try {
    updateLeadTags({
      rowIndex: rowIndex,
      sheetName: targetSheetName,
      manualTags: lead.manualTags || []
    });
  } catch (tagError) {
    console.warn("Auto tag refresh failed after promo redemption:", tagError);
  }

  return {
    success: true,
    rowIndex: rowIndex
  };
}

/*************************************************************
 *  updateB12Offered - Mark B12 as offered
 *************************************************************/
function updateB12Offered(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);
  const b12Offered = Boolean(payload.b12Offered);

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const targetSheetName = payload.sheetName || CUSTOMERS_SHEET_NAME;
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  ensureColumnExists(sheet, COL.B12_OFFERED);
  sheet.getRange(rowIndex, COL.B12_OFFERED).setValue(b12Offered ? "TRUE" : "FALSE");

  const lead = getSelectedLeadFromRow(rowIndex, targetSheetName);
  injectLeadIntoSidebar(lead);

  return {
    success: true,
    rowIndex: rowIndex,
    b12Offered: b12Offered
  };
}

/*************************************************************
 *  updateB12Redeemed - Mark B12 as redeemed
 *************************************************************/
function updateB12Redeemed(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);
  const b12Redeemed = Boolean(payload.b12Redeemed);

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const targetSheetName = payload.sheetName || CUSTOMERS_SHEET_NAME;
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  ensureColumnExists(sheet, COL.B12_REDEEMED);
  sheet.getRange(rowIndex, COL.B12_REDEEMED).setValue(b12Redeemed ? "TRUE" : "FALSE");

  // If B12 is redeemed, also mark as offered
  if (b12Redeemed) {
    sheet.getRange(rowIndex, COL.B12_OFFERED).setValue("TRUE");
  }

  const lead = getSelectedLeadFromRow(rowIndex, targetSheetName);
  injectLeadIntoSidebar(lead);

  return {
    success: true,
    rowIndex: rowIndex,
    b12Redeemed: b12Redeemed
  };
}

/*************************************************************
 *  logNewMembership - Record membership details and promo usage
 *************************************************************/
function logNewMembership(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);
  const targetSheetName = payload.sheetName || CUSTOMERS_SHEET_NAME;
  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  const cleanedName = cleanValue(payload.name || "");
  const cleanedEmail = cleanValue(payload.email || "");
  const cleanedPhone = cleanValue(payload.phone || "");
  const membershipType = (payload.membershipType || "Other").toString().trim() || "Other";
  const promoSelection = (payload.promoSelection || "").toString().trim();
  const promoUsed = Boolean(payload.promoUsed);
  const notes = (payload.notes || "").toString().trim();
  const tz = Session.getScriptTimeZone();

  let consultationDateValue = "";
  if (payload.consultationDate) {
    const parsed = new Date(payload.consultationDate);
    if (!isNaN(parsed.getTime())) {
      consultationDateValue = parsed;
    } else {
      throw new Error("Invalid consultation date provided.");
    }
  }

  ensureColumnExists(sheet, COL.NAME);
  ensureColumnExists(sheet, COL.EMAIL);
  ensureColumnExists(sheet, COL.PHONE);
  ensureColumnExists(sheet, COL.CONSULTATION_DATE);
  ensureColumnExists(sheet, COL.BOOKED_CONSULTATION);
  ensureColumnExists(sheet, COL.JOINED_MEMBERSHIP);
  ensureColumnExists(sheet, COL.MEMBERSHIP_TYPE);
  ensureColumnExists(sheet, COL.MEMBERSHIP_START_DATE);
  ensureColumnExists(sheet, COL.LEAD_STATUS);

  if (cleanedName) {
    sheet.getRange(rowIndex, COL.NAME).setValue(cleanedName);
  }
  if (cleanedEmail) {
    sheet.getRange(rowIndex, COL.EMAIL).setValue(cleanedEmail);
  }
  if (cleanedPhone) {
    sheet.getRange(rowIndex, COL.PHONE).setValue(cleanedPhone);
  }

  sheet.getRange(rowIndex, COL.CONSULTATION_DATE).setValue(consultationDateValue || "");
  sheet.getRange(rowIndex, COL.BOOKED_CONSULTATION).setValue("TRUE");
  sheet.getRange(rowIndex, COL.JOINED_MEMBERSHIP).setValue("TRUE");
  sheet.getRange(rowIndex, COL.MEMBERSHIP_TYPE).setValue(membershipType);
  sheet.getRange(rowIndex, COL.MEMBERSHIP_START_DATE).setValue(new Date());
  sheet.getRange(rowIndex, COL.LEAD_STATUS).setValue("Converted");

  if (promoUsed) {
    ensureColumnExists(sheet, COL.PROMO_REDEEMED);
    ensureColumnExists(sheet, COL.PROMO_DESCRIPTION);
    ensureColumnExists(sheet, COL.B12_OFFERED);
    ensureColumnExists(sheet, COL.B12_REDEEMED);

    const normalizedPromo = promoSelection.toLowerCase();
    const includesB12 = normalizedPromo === "free b12 shot" || normalizedPromo === "$25 off first month & b12 shot";
    const includesTwentyFiveOff = normalizedPromo === "$25 off first month" || normalizedPromo === "$25 off first month & b12 shot";

    if (promoSelection) {
      sheet.getRange(rowIndex, COL.PROMO_DESCRIPTION).setValue(promoSelection);
    }
    if (includesTwentyFiveOff) {
      sheet.getRange(rowIndex, COL.PROMO_REDEEMED).setValue("TRUE");
    }
    if (includesB12) {
      sheet.getRange(rowIndex, COL.B12_OFFERED).setValue("TRUE");
      sheet.getRange(rowIndex, COL.B12_REDEEMED).setValue("TRUE");
    }
  }

  const summaryParts = ["Membership Logged", "Type: " + membershipType];
  if (consultationDateValue instanceof Date) {
    summaryParts.push("Consultation: " + Utilities.formatDate(consultationDateValue, tz, "MMM d, yyyy"));
  }
  if (promoUsed && promoSelection) {
    summaryParts.push("Promo: " + promoSelection);
  }

  let noteToAppend = summaryParts.join(" | ");
  if (notes) {
    noteToAppend += "\n" + notes;
  }

  appendLeadNotes({
    rowIndex: rowIndex,
    note: noteToAppend,
    sheetName: targetSheetName
  });

  const updatedLead = getSelectedLeadFromRow(rowIndex, targetSheetName);
  injectLeadIntoSidebar(updatedLead);

  try {
    updateLeadTags({
      rowIndex: rowIndex,
      sheetName: targetSheetName,
      manualTags: updatedLead.manualTags || []
    });
  } catch (tagError) {
    console.warn("Auto tag refresh failed after logging membership:", tagError);
  }

  return {
    success: true,
    lead: updatedLead
  };
}

function createNewPatient(payload) {
  if (!payload) {
    throw new Error("Payload is required.");
  }

  const name = cleanValue(payload.name || "");
  if (!name) {
    throw new Error("Patient name is required.");
  }

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const targetSheetName = payload.sheetName || CUSTOMERS_SHEET_NAME;
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    throw new Error("Sheet '" + targetSheetName + "' not found.");
  }

  ensureColumnExists(sheet, COL.CLOSED_FROM);

  const insertIndex = HEADER_ROW_INDEX + 1;
  sheet.insertRowsAfter(HEADER_ROW_INDEX, 1);
  sheet.setRowHeight(insertIndex, 90);

  const now = new Date();
  const rowValues = Array(COL.CLOSED_FROM).fill("");
  rowValues[COL.TIMESTAMP - 1] = now;
  rowValues[COL.LEAD_ID - 1] = payload.leadId || generateLeadId(sheet.getLastRow() + 1);
  rowValues[COL.SOURCE - 1] = cleanValue(payload.source || "Manual Entry");
  rowValues[COL.GENDER - 1] = cleanValue(payload.gender || "");
  rowValues[COL.NAME - 1] = name;
  rowValues[COL.EMAIL - 1] = cleanValue(payload.email || "");
  rowValues[COL.PHONE - 1] = cleanValue(payload.phone || "");
  rowValues[COL.AGE_RANGE - 1] = cleanValue(payload.age || "");
  rowValues[COL.LEAD_STATUS - 1] = cleanValue(payload.leadStatus || "New");
  rowValues[COL.FOLLOW_UP_PRIORITY - 1] = cleanValue(payload.followUpPriority || "Medium");
  rowValues[COL.LEAD_TYPE - 1] = cleanValue(payload.leadType || "Manual Entry");
  rowValues[COL.PROMO_CODE - 1] = cleanValue(payload.promoCode || "");
  rowValues[COL.PROMO_DESCRIPTION - 1] = cleanValue(payload.promoDescription || "");

  if (payload.scorePercent !== undefined && payload.scorePercent !== null && payload.scorePercent !== "") {
    rowValues[COL.BIO_BALANCE_SCORE - 1] = Number(payload.scorePercent) || 0;
  }

  applyQuizNotesToRow(rowValues, payload.quizNotes);

  if (Array.isArray(payload.quizNotes) && payload.quizNotes.length) {
    try {
      rowValues[COL.RAW_ANSWERS - 1] = JSON.stringify(payload.quizNotes);
    } catch (err) {
      rowValues[COL.RAW_ANSWERS - 1] = "";
    }
  }

  sheet.getRange(insertIndex, 1, 1, rowValues.length).setValues([rowValues.map(cleanValue)]);

  const manualTagInput = payload.manualTags !== undefined ? payload.manualTags : payload.tags;
  const manualTags = parseTagsInput(manualTagInput);
  let tagResult = null;
  try {
    tagResult = updateLeadTags({
      rowIndex: insertIndex,
      sheetName: targetSheetName,
      manualTags: manualTags
    });
  } catch (tagError) {
    console.warn("Unable to set initial tags for new patient:", tagError);
  }

  let lead = (tagResult && tagResult.lead) || getSelectedLeadFromRow(insertIndex, targetSheetName);
  injectLeadIntoSidebar(lead);

  if (payload.generalNotes) {
    appendLeadNotes({
      rowIndex: insertIndex,
      note: payload.generalNotes,
      sheetName: targetSheetName
    });
    lead = getSelectedLeadFromRow(insertIndex, targetSheetName);
  }

  let reportResult = null;
  if (payload.generateReport) {
    if (lead.email) {
      try {
        reportResult = generateReportForLead({
          rowIndex: insertIndex,
          sheetName: targetSheetName,
          lead: lead,
          sendEmail: true
        });
      } catch (err) {
        reportResult = {
          success: false,
          error: err.message
        };
      }
    } else {
      reportResult = {
        success: false,
        error: "Email is required to send the report."
      };
    }
  }

  lead = getSelectedLeadFromRow(insertIndex, targetSheetName);
  injectLeadIntoSidebar(lead);

  return {
    success: true,
    rowIndex: insertIndex,
    lead: lead,
    report: reportResult
  };
}

function generateReportForLead(options) {
  if (!options || !options.rowIndex) {
    throw new Error("Row index is required.");
  }

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheetName = options.sheetName || CUSTOMERS_SHEET_NAME;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("Sheet '" + sheetName + "' not found.");
  }

  const lead = options.lead || getSelectedLeadFromRow(options.rowIndex, sheetName);

  // If staff selected a promo override, use it
  let selectedPromo = options.promoCode || "";
  if (selectedPromo) {
    // Apply the flat $25 / VIP rules
    options.promoDescription = autoPromoDescription(selectedPromo);
  }

  if (!selectedPromo) {
    selectedPromo = lead.promoCode || "";
  } else {
    lead.promoCode = selectedPromo;
    lead.promoDescription = options.promoDescription || lead.promoDescription;
  }

  const prompt = buildGeminiPrompt({
    name: lead.name,
    gender: lead.gender,
    ageRange: lead.ageRange,
    scorePercent: Number(lead.scorePercent || 0),
    promoCode: selectedPromo,
    source: lead.source || "Manual Entry",
    sleepQuality: lead.sleepQuality,
    energyLevels: lead.energyLevels,
    mentalClarity: lead.mentalClarity,
    stressTolerance: lead.stressTolerance,
    activityLevel: lead.activityLevel,
    bodyChanges: lead.bodyChanges,
    libido: lead.libido,
    tempSensitivityF: lead.tempSensitivityF,
    cycleStatusF: lead.cycleStatusF,
    pmsMoodF: lead.pmsMoodF,
    bloatingF: lead.bloatingF,
    strengthRecoveryM: lead.strengthRecoveryM,
    morningDriveM: lead.morningDriveM,
    muscleMassM: lead.muscleMassM,
    sexualPerformanceM: lead.sexualPerformanceM,
    symptoms: lead.symptoms || [],
    notes: lead.notes || ""
  });

  let reportText;
  try {
    reportText = callGemini(prompt);
  } catch (err) {
    reportText = "<p>We were unable to auto-generate your report. A provider will follow up shortly.</p>";
  }

  const onscreenReportText = stripOnscreenPromotions(reportText);
  const reportHtml = buildReportHtml({
    name: lead.name || "",
    gender: lead.gender || "",
    promoCode: selectedPromo || "",
    reportText: onscreenReportText,
    scorePercent: Number(lead.scorePercent || 0)
  });

  let emailed = false;
  if (options.sendEmail !== false && lead.email) {
    const emailHtml = composeReportEmailHtml({
      name: lead.name || "",
      gender: lead.gender || "",
      ageRange: lead.ageRange || "",
      promoCode: selectedPromo || "",
      reportText: reportText
    });

    MailApp.sendEmail({
      to: lead.email,
      subject: "Your Personalized Bio-Balance Report – White Tank Wellness",
      htmlBody: emailHtml
    });

    emailed = true;
  }

  if (lead.leadId) {
    updateStatus(lead.leadId, 6, "Manual report generated");
  }

  const noteSummary = emailed
    ? "Manual report generated and emailed to " + lead.email
    : "Manual report generated (email not sent)";

  appendLeadNotes({
    rowIndex: options.rowIndex,
    note: noteSummary,
    sheetName: sheetName
  });

  return {
    success: true,
    reportHtml: reportHtml,
    emailed: emailed,
    leadId: lead.leadId || "",
    rowIndex: options.rowIndex
  };
}

/*************************************************************
 *  recalculateLeadScores - Batch update sales scores & buckets
 *************************************************************/
function recalculateLeadScores() {
  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet 'Submissions' not found.");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW_INDEX) {
    return { success: true, message: "No leads to process." };
  }

  ensureColumnExists(sheet, COL.RETARGET_BUCKET);

  let updatedCount = 0;

  for (let row = HEADER_ROW_INDEX + 1; row <= lastRow; row++) {
    try {
      const lead = getSelectedLeadFromRow(row);
      
      // Calculate sales score
      const symptomData = {
        ageRange: lead.ageRange,
        sleepQuality: lead.sleepQuality,
        energyLevels: lead.energyLevels,
        mentalClarity: lead.mentalClarity,
        stressTolerance: lead.stressTolerance,
        email: lead.email,
        phone: lead.phone,
        name: lead.name
      };
      const salesScore = calculateSalesLeadScore(lead.scorePercent, symptomData);
      
      // Calculate retarget bucket
      const retargetBucket = autoRetargetBucket(lead);

      // Update cells
      sheet.getRange(row, COL.LEAD_SCORE_SALES).setValue(salesScore);
      sheet.getRange(row, COL.RETARGET_BUCKET).setValue(retargetBucket);
      
      updatedCount++;
    } catch (err) {
      console.error("Error processing row " + row + ": " + err);
    }
  }

  return {
    success: true,
    message: "Updated " + updatedCount + " leads."
  };
}

/*************************************************************
 *  logOutreachAttempt - Log outreach with method type
 *************************************************************/
function logOutreachAttempt(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);
  const method = (payload.method || "Contact").trim();
  const note = (payload.note || "").trim();

  const formattedNote = method + ": " + (note || "Attempted contact");

  return appendLeadNotes({
    rowIndex: rowIndex,
    note: formattedNote,
    isOutreach: true
  });
}

/*************************************************************
 *  SEARCH FUNCTIONALITY
 *************************************************************/
function searchLeadByPromoOrEmail(searchTerm) {
  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CUSTOMERS_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet '" + CUSTOMERS_SHEET_NAME + "' not found.");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW_INDEX) {
    return [];
  }

  const lastColumn = sheet.getLastColumn() || COL.PROMO_CODE;
  const data = sheet.getRange(HEADER_ROW_INDEX + 1, 1, lastRow - HEADER_ROW_INDEX, lastColumn).getValues();
  
  const results = [];
  const searchLower = searchTerm.toLowerCase().trim();
  
  for (let i = 0; i < data.length; i++) {
    const email = (data[i][COL.EMAIL - 1] || '').toString().toLowerCase();
    const promo = (data[i][COL.PROMO_CODE - 1] || '').toString().toLowerCase();
    const leadId = (data[i][COL.LEAD_ID - 1] || '').toString().toLowerCase();
    const name = (data[i][COL.NAME - 1] || '').toString().toLowerCase();
    const phone = (data[i][COL.PHONE - 1] || '').toString().toLowerCase();
    
    if (email.includes(searchLower) || 
        promo.includes(searchLower) || 
        leadId.includes(searchLower) ||
        name.includes(searchLower) ||
        phone.includes(searchLower)) {
      results.push({
        row: i + HEADER_ROW_INDEX + 1,
        name: data[i][COL.NAME - 1] || '',
        email: data[i][COL.EMAIL - 1] || '',
        phone: data[i][COL.PHONE - 1] || '',
        promoCode: data[i][COL.PROMO_CODE - 1] || '',
        leadId: data[i][COL.LEAD_ID - 1] || '',
        leadStatus: data[i][COL.LEAD_STATUS - 1] || 'New',
        scorePercent: data[i][COL.BIO_BALANCE_SCORE - 1] || 0
      });
    }
  }
  
  return results;
}

function getLeadByRow(row) {
  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CUSTOMERS_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet '" + CUSTOMERS_SHEET_NAME + "' not found.");
  }

  // Set active range to the selected row
  sheet.setActiveRange(sheet.getRange(row, 1));

  const lead = getSelectedLeadFromRow(row, CUSTOMERS_SHEET_NAME);
  injectLeadIntoSidebar(lead);
  return lead;
}

function filterLeads(query) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  query = query.toString().trim().toLowerCase();

  const results = data
    .map((row, i) => {
      return {
        rowIndex: i + 2,
        timestamp: row[0],
        leadId: row[1],
        source: row[2],
        name: row[8],
        email: row[9],
        phone: row[10],
        ageRange: row[11],
        score: row[27],
        tags: row[50]
      };
    })
    .filter(lead => {
      return (
        (lead.name || "").toString().toLowerCase().includes(query) ||
        (lead.email || "").toString().toLowerCase().includes(query) ||
        (lead.phone || "").toString().toLowerCase().includes(query) ||
        (lead.tags || "").toString().toLowerCase().includes(query) ||
        (lead.leadId || "").toString().toLowerCase().includes(query)
      );
    });

  return results;
}

function sanitizeFilterArray(list) {
  if (!Array.isArray(list)) {
    return [];
  }
  return list
    .map(cleanValue)
    .map(function(value) {
      return value === null || value === undefined ? "" : String(value).trim();
    })
    .filter(function(value) {
      return Boolean(value);
    });
}

function parseFilterNumber(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const numeric = Number(value);
  return isNaN(numeric) ? null : numeric;
}

function parseFilterBoolean(value) {
  if (value === true || value === false) {
    return value;
  }
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "true" || normalized === "yes" || normalized === "1") {
      return true;
    }
    if (normalized === "false" || normalized === "no" || normalized === "0") {
      return false;
    }
  }
  if (typeof value === "number") {
    return value === 1;
  }
  return false;
}

function sanitizePipelineStageList(list) {
  const sanitized = [];
  if (!Array.isArray(list)) {
    return sanitized;
  }
  list.forEach(function(entry) {
    const direct = entry && PIPELINE_STAGE_KEYS.indexOf(entry) >= 0 ? entry : null;
    if (direct && sanitized.indexOf(direct) === -1) {
      sanitized.push(direct);
      return;
    }
    const normalized = normalizeFilterString(entry);
    if (!normalized) {
      return;
    }
    const match = PIPELINE_STAGE_KEYS.find(function(stage) {
      return normalizeFilterString(stage) === normalized;
    });
    if (match && sanitized.indexOf(match) === -1) {
      sanitized.push(match);
    }
  });
  return sanitized;
}

function sanitizeAdvancedFilterPayload(filters) {
  const source = filters && typeof filters === "object" ? filters : {};
  const sanitized = {};

  sanitized.ageRanges = sanitizeFilterArray(source.ageRanges);
  sanitized.ageRangeSet = new Set(sanitized.ageRanges.map(normalizeFilterString));

  sanitized.genders = sanitizeFilterArray(source.genders);
  sanitized.genderSet = new Set(sanitized.genders.map(normalizeFilterString));

  sanitized.symptomGroups = sanitizeFilterArray(source.symptomGroups)
    .map(function(entry) {
      const normalized = normalizeFilterString(entry);
      return SYMPTOM_FIELD_LOOKUP[normalized] || entry;
    })
    .filter(Boolean);
  sanitized.symptomGroups = Array.from(new Set(sanitized.symptomGroups));
  sanitized.symptomSeverity = Number(source.symptomSeverity || 0);
  if (!isFinite(sanitized.symptomSeverity) || sanitized.symptomSeverity < 0) {
    sanitized.symptomSeverity = 0;
  }
  if (sanitized.symptomSeverity > 4) {
    sanitized.symptomSeverity = 4;
  }

  let minScore = parseFilterNumber(source.bioScoreMin);
  let maxScore = parseFilterNumber(source.bioScoreMax);
  function clampScore(value) {
    if (value === null) return null;
    if (value < 0) return 0;
    if (value > 100) return 100;
    return value;
  }
  minScore = clampScore(minScore);
  maxScore = clampScore(maxScore);
  if (minScore !== null && maxScore !== null && minScore > maxScore) {
    const temp = minScore;
    minScore = maxScore;
    maxScore = temp;
  }
  sanitized.bioScoreMin = minScore;
  sanitized.bioScoreMax = maxScore;

  sanitized.leadTypes = sanitizeFilterArray(source.leadTypes);
  sanitized.leadTypeSet = new Set(sanitized.leadTypes.map(normalizeFilterString));
  sanitized.leadTypeKeywordLower = normalizeFilterString(source.leadTypeKeyword);

  sanitized.booking = source.booking === "booked" || source.booking === "notBooked" ? source.booking : "any";
  sanitized.promoStatus = source.promoStatus === "hasPromo" || source.promoStatus === "noPromo" ? source.promoStatus : "any";
  sanitized.promoTextLower = normalizeFilterString(source.promoText);

  sanitized.b12Only = parseFilterBoolean(source.b12Only);
  sanitized.promoRedeemedOnly = parseFilterBoolean(source.promoRedeemedOnly);

  sanitized.pipelineStages = sanitizePipelineStageList(source.pipelineStages);
  sanitized.pipelineStageSet = new Set(sanitized.pipelineStages);

  sanitized.tagsLower = normalizeFilterString(source.tags);

  return sanitized;
}

function hasActiveAdvancedFilters(filters) {
  if (!filters) {
    return false;
  }
  if (filters.ageRanges.length || filters.genders.length || filters.symptomGroups.length || filters.leadTypes.length || filters.pipelineStages.length) {
    return true;
  }
  if (filters.bioScoreMin !== null || filters.bioScoreMax !== null) {
    return true;
  }
  if (filters.leadTypeKeywordLower) {
    return true;
  }
  if (filters.booking !== "any" || filters.promoStatus !== "any") {
    return true;
  }
  if (filters.promoTextLower) {
    return true;
  }
  if (filters.b12Only || filters.promoRedeemedOnly) {
    return true;
  }
  if (filters.tagsLower) {
    return true;
  }
  return false;
}

function parseLeadDateValue(value) {
  if (!value) {
    return null;
  }
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function determinePipelineStageForLead(lead) {
  if (!lead) {
    return "quizComplete";
  }
  const status = normalizeFilterString(lead.leadStatus);
  if (lead.joinedMembership || status === "converted") {
    return "member";
  }
  const booked = lead.bookedConsultation || status === "booked";
  if (!booked) {
    return "quizComplete";
  }
  const consultDate = parseLeadDateValue(lead.consultationDateRaw || lead.consultationDateIso || lead.consultationDate);
  if (consultDate && consultDate.getTime() <= Date.now()) {
    return "consultNoJoin";
  }
  return "bookedConsult";
}

function leadHasPromo(lead) {
  if (!lead) {
    return false;
  }
  const promoCode = lead.promoCode ? lead.promoCode.toString().trim() : "";
  const promoDescription = lead.promoDescription ? lead.promoDescription.toString().trim() : "";
  return Boolean(promoCode || promoDescription);
}

function leadPromoTextIncludes(lead, searchLower) {
  if (!searchLower) {
    return true;
  }
  const textParts = [];
  if (lead && lead.promoCode) {
    textParts.push(String(lead.promoCode));
  }
  if (lead && lead.promoDescription) {
    textParts.push(String(lead.promoDescription));
  }
  const combined = textParts.join(" ").toLowerCase();
  return combined.includes(searchLower);
}

function matchesAdvancedFilters(lead, filters, pipelineStage) {
  if (filters.ageRangeSet.size) {
    const ageNormalized = normalizeFilterString(lead.ageRange);
    if (!filters.ageRangeSet.has(ageNormalized)) {
      return false;
    }
  }
  if (filters.genderSet.size) {
    const genderNormalized = normalizeFilterString(lead.gender);
    if (!filters.genderSet.has(genderNormalized)) {
      return false;
    }
  }
  if (filters.leadTypeSet.size) {
    const leadTypeNormalized = normalizeFilterString(lead.leadType);
    if (!filters.leadTypeSet.has(leadTypeNormalized)) {
      return false;
    }
  }
  if (filters.leadTypeKeywordLower) {
    const leadTypeNormalized = normalizeFilterString(lead.leadType);
    if (!leadTypeNormalized.includes(filters.leadTypeKeywordLower)) {
      return false;
    }
  }
  if (filters.bioScoreMin !== null && Number(lead.scorePercent || 0) < filters.bioScoreMin) {
    return false;
  }
  if (filters.bioScoreMax !== null && Number(lead.scorePercent || 0) > filters.bioScoreMax) {
    return false;
  }
  if (filters.booking === "booked" && !(lead.bookedConsultation || normalizeFilterString(lead.leadStatus) === "booked")) {
    return false;
  }
  if (filters.booking === "notBooked" && (lead.bookedConsultation || normalizeFilterString(lead.leadStatus) === "booked")) {
    return false;
  }
  if (filters.promoStatus === "hasPromo" && !leadHasPromo(lead)) {
    return false;
  }
  if (filters.promoStatus === "noPromo" && leadHasPromo(lead)) {
    return false;
  }
  if (filters.b12Only && !lead.b12Offered) {
    return false;
  }
  if (filters.promoRedeemedOnly && !lead.promoRedeemed) {
    return false;
  }
  if (filters.promoTextLower && !leadPromoTextIncludes(lead, filters.promoTextLower)) {
    return false;
  }
  if (filters.pipelineStageSet.size && !filters.pipelineStageSet.has(pipelineStage)) {
    return false;
  }
  if (filters.symptomGroups.length) {
    const rawAnswers = parseLeadRawAnswers(lead.rawAnswers);
    const matchesSymptom = filters.symptomGroups.some(function(fieldKey) {
      const severity = getSymptomSeverityForLead(lead, rawAnswers, fieldKey);
      return severity >= filters.symptomSeverity;
    });
    if (!matchesSymptom) {
      return false;
    }
  }
  if (filters.tagsLower) {
    const leadTags = lead.tags || "";
    const leadTagsLower = String(leadTags).toLowerCase();
    const searchTerms = filters.tagsLower.split(',').map(function(term) { return term.trim(); }).filter(Boolean);
    const matchesAllTerms = searchTerms.every(function(term) {
      return leadTagsLower.includes(term);
    });
    if (!matchesAllTerms) {
      return false;
    }
  }
  return true;
}

function formatLeadForFilterResult(lead, pipelineStage) {
  return {
    row: lead.rowIndex,
    name: lead.name || "",
    email: lead.email || "",
    phone: lead.phone || "",
    ageRange: lead.ageRange || "",
    gender: lead.gender || "",
    scorePercent: lead.scorePercent || 0,
    leadType: lead.leadType || "",
    promoCode: lead.promoCode || "",
    promoDescription: lead.promoDescription || "",
    pipelineStage: pipelineStage,
    bookedConsultation: Boolean(lead.bookedConsultation),
    joinedMembership: Boolean(lead.joinedMembership),
    leadStatus: lead.leadStatus || ""
  };
}

function filterLeadsAdvanced(filters) {
  const sanitized = sanitizeAdvancedFilterPayload(filters);
  if (!hasActiveAdvancedFilters(sanitized)) {
    throw new Error("Select at least one filter before applying.");
  }

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet 'Submissions' not found.");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW_INDEX) {
    return [];
  }

  const lastColumn = sheet.getLastColumn() || COL.CLOSED_FROM;
  const data = sheet.getRange(HEADER_ROW_INDEX + 1, 1, lastRow - HEADER_ROW_INDEX, lastColumn).getValues();
  const results = [];

  data.forEach(function(rowValues, index) {
    const rowIndex = HEADER_ROW_INDEX + 1 + index;
    const lead = buildLeadObjectFromRow(rowIndex, rowValues);
    const pipelineStage = determinePipelineStageForLead(lead);
    if (!matchesAdvancedFilters(lead, sanitized, pipelineStage)) {
      return;
    }
    results.push(formatLeadForFilterResult(lead, pipelineStage));
  });

  results.sort(function(a, b) {
    return (Number(b.scorePercent) || 0) - (Number(a.scorePercent) || 0);
  });

  return results;
}

/*************************************************************
 *  mark72hrFollowUpSent - Mark 72hr follow-up as sent
 *************************************************************/
function mark72hrFollowUpSent(payload) {
  if (!payload || !payload.rowIndex) {
    throw new Error("A valid row selection is required.");
  }

  const rowIndex = Number(payload.rowIndex);

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet 'Submissions' not found.");
  }

  ensureColumnExists(sheet, COL.FOLLOW_UP_72HR_TIMESTAMP);
  
  sheet.getRange(rowIndex, COL.FOLLOW_UP_SENT).setValue("TRUE");
  sheet.getRange(rowIndex, COL.FOLLOW_UP_72HR_TIMESTAMP).setValue(new Date());

  const lead = getSelectedLeadFromRow(rowIndex);
  injectLeadIntoSidebar(lead);

  return {
    success: true,
    rowIndex: rowIndex
  };
}

/*************************************************************
 *  REPORT CLOSE HANDLER & FOLLOW-UP AUTOMATION
 *************************************************************/
function processReportClose(payload) {
  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);

  const leadId = (payload.leadId || "").toString();
  const leadType = payload.leadType || "Cold";
  const name = payload.name || payload.contactInfo?.name || "";
  const email = payload.email || payload.contactInfo?.email || "";
  const phone = payload.phone || payload.contactInfo?.phone || "";
  const scorePercent = Number(payload.scorePercent ?? payload.bioBalanceScore ?? 0) || 0;
  const promoCode = payload.promoCode || "";
  const sourceUrl = payload.pageUrl || payload.utmLandingPage || "";
  const referrer = payload.referrer || "";
  const source = payload.source || "Bio-Balance Report Close";
  const now = new Date();

  let rowIndex = null;
  if (sheet) {
    ensureColumnExists(sheet, COL.LEAD_TYPE);
    if (leadId) {
      rowIndex = findLeadRowByLeadId(sheet, leadId);
    }
    if (!rowIndex && email) {
      rowIndex = findLeadRowByEmail(sheet, email);
    }
    if (rowIndex) {
      sheet.getRange(rowIndex, COL.LEAD_TYPE).setValue(leadType);
    }
  }

  const triggerInfo = create72hrFollowUpTrigger({
    leadId: leadId || (email ? "EMAIL-" + email : ""),
    email: email,
    name: name,
    leadType: leadType,
    scorePercent: scorePercent,
    promoCode: promoCode,
    sourceUrl: sourceUrl,
    rowIndex: rowIndex
  });

  appendCloseAudit({
    timestamp: now,
    action: "report_closed",
    leadId: leadId || "(missing)",
    email: email,
    name: name,
    leadType: leadType,
    scorePercent: scorePercent,
    promoCode: promoCode,
    pageUrl: sourceUrl,
    referrer: referrer,
    triggerId: triggerInfo ? triggerInfo.triggerId : "",
    scheduledFor: triggerInfo ? triggerInfo.scheduledFor : "",
    details: rowIndex ? "Matched lead in sheet" : "Lead not found by id/email",
    source: source
  });

  if (email) {
    sendCloseAcknowledgementEmail({
      email: email,
      name: name,
      scorePercent: scorePercent,
      promoCode: promoCode,
      leadType: leadType
    });
  }

  trackCloseEvents({
    leadId: leadId || email || Utilities.getUuid(),
    leadType: leadType,
    scorePercent: scorePercent,
    promoCode: promoCode,
    sourceUrl: sourceUrl || "https://www.whitetankwellness.com",
    referrer: referrer,
    email: email,
    phone: phone
  });

  return {
    status: "close-recorded",
    leadId: leadId || "",
    rowIndex: rowIndex,
    followUpScheduledFor: triggerInfo ? triggerInfo.scheduledFor : null,
    triggerId: triggerInfo ? triggerInfo.triggerId : null
  };
}

function appendCloseAudit(entry) {
  try {
    const sheet = getCloseAuditSheet();
    sheet.appendRow([
      entry.timestamp || new Date(),
      entry.leadId || "",
      entry.email || "",
      entry.name || "",
      entry.leadType || "",
      entry.scorePercent || "",
      entry.promoCode || "",
      entry.action || "",
      entry.details || "",
      entry.pageUrl || "",
      entry.referrer || "",
      entry.triggerId || "",
      entry.scheduledFor || ""
    ]);
  } catch (err) {
    console.error("appendCloseAudit error:", err);
  }
}

function getCloseAuditSheet() {
  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CLOSE_AUDIT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CLOSE_AUDIT_SHEET_NAME);
    sheet.appendRow([
      "Timestamp",
      "Lead ID",
      "Email",
      "Name",
      "Lead Type",
      "Score %",
      "Promo Code",
      "Action",
      "Details",
      "Page URL",
      "Referrer",
      "Trigger ID",
      "Follow-Up Scheduled For"
    ]);
  }
  return sheet;
}

function getScriptProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || "";
}

function trackCloseEvents(options) {
  try {
    sendGa4Event("report_close", {
      lead_id: options.leadId || "",
      lead_type: options.leadType || "",
      score_percent: options.scorePercent || 0,
      promo_code: options.promoCode || "",
      source_url: options.sourceUrl || "",
      referrer: options.referrer || ""
    }, options.leadId || Utilities.getUuid());
  } catch (err) {
    console.error("GA4 close tracking error:", err);
  }

  try {
    sendMetaEvent({
      leadId: options.leadId,
      leadType: options.leadType,
      scorePercent: options.scorePercent,
      promoCode: options.promoCode,
      email: options.email,
      phone: options.phone,
      sourceUrl: options.sourceUrl
    });
  } catch (err) {
    console.error("Meta close tracking error:", err);
  }
}

function sendGa4Event(eventName, params, clientId) {
  const measurementId = getScriptProp(GA4_MEASUREMENT_ID_PROP);
  const apiSecret = getScriptProp(GA4_API_SECRET_PROP);
  if (!measurementId || !apiSecret) {
    return { success: false, message: "GA4 not configured" };
  }

  const url = "https://www.google-analytics.com/mp/collect?measurement_id=" + measurementId + "&api_secret=" + apiSecret;
  const body = {
    client_id: clientId || Utilities.getUuid(),
    events: [{
      name: eventName,
      params: params
    }]
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  return {
    success: response.getResponseCode() >= 200 && response.getResponseCode() < 300,
    status: response.getResponseCode()
  };
}

function sendMetaEvent(options) {
  const pixelId = getScriptProp(META_PIXEL_ID_PROP);
  const accessToken = getScriptProp(META_ACCESS_TOKEN_PROP);
  if (!pixelId || !accessToken) {
    return { success: false, message: "Meta not configured" };
  }

  const url = "https://graph.facebook.com/v18.0/" + pixelId + "/events?access_token=" + accessToken;

  const userData = {};
  const emailHash = hashForMeta(options.email);
  if (emailHash) userData.em = emailHash;
  const phoneHash = hashForMeta(options.phone);
  if (phoneHash) userData.ph = phoneHash;
  const externalHash = hashForMeta(options.leadId);
  if (externalHash) userData.external_id = externalHash;

  const payload = {
    data: [{
      event_name: "ReportClose",
      event_time: Math.floor(Date.now() / 1000),
      event_id: (options.leadId || Utilities.getUuid()) + "-close",
      action_source: "website",
      event_source_url: options.sourceUrl || "https://www.whitetankwellness.com",
      user_data: userData,
      custom_data: {
        lead_type: options.leadType || "",
        score_percent: options.scorePercent || 0,
        promo_code: options.promoCode || ""
      }
    }]
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  return {
    success: response.getResponseCode() >= 200 && response.getResponseCode() < 300,
    status: response.getResponseCode()
  };
}

function hashForMeta(value) {
  if (!value) return "";
  const normalized = value.toString().trim().toLowerCase();
  if (!normalized) return "";

  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, normalized);
  const hex = raw.map(function(b) {
    const v = (b + 256) % 256;
    return (v < 16 ? "0" : "") + v.toString(16);
  }).join("");
  return hex;
}

function create72hrFollowUpTrigger(details) {
  if (!details || !details.email) return null;

  const leadId = details.leadId || "EMAIL-" + details.email;
  const props = PropertiesService.getScriptProperties();

  const existingJson = props.getProperty(FOLLOW_UP_PROPERTY_PREFIX + leadId);
  if (existingJson) {
    try {
      const existing = JSON.parse(existingJson);
      cleanupFollowUpArtifacts(leadId, existing.triggerId);
    } catch (err) {
      cleanupFollowUpArtifacts(leadId, null);
    }
  }

  const trigger = ScriptApp.newTrigger("handle72hrFollowUp")
    .timeBased()
    .after(72 * 60 * 60 * 1000)
    .create();

  const scheduledFor = new Date(Date.now() + 72 * 60 * 60 * 1000);
  const record = {
    leadId: leadId,
    email: details.email,
    name: details.name || "",
    leadType: details.leadType || "Cold",
    scorePercent: Number(details.scorePercent || 0),
    promoCode: details.promoCode || "",
    sourceUrl: details.sourceUrl || "",
    rowIndex: details.rowIndex || null,
    triggerId: trigger.getUniqueId(),
    createdAt: new Date().toISOString(),
    scheduledFor: scheduledFor.toISOString()
  };

  props.setProperty(FOLLOW_UP_PROPERTY_PREFIX + leadId, JSON.stringify(record));
  props.setProperty(TRIGGER_LOOKUP_PREFIX + trigger.getUniqueId(), leadId);

  appendCloseAudit({
    timestamp: new Date(),
    action: "followup_scheduled",
    leadId: leadId,
    email: details.email,
    name: details.name || "",
    leadType: details.leadType || "Cold",
    scorePercent: details.scorePercent || "",
    promoCode: details.promoCode || "",
    pageUrl: details.sourceUrl || "",
    referrer: "",
    triggerId: trigger.getUniqueId(),
    scheduledFor: scheduledFor,
    details: "72-hour follow-up scheduled"
  });

  return {
    triggerId: trigger.getUniqueId(),
    scheduledFor: scheduledFor
  };
}

function handle72hrFollowUp(e) {
  try {
    const triggerId = e && e.triggerUid ? e.triggerUid : null;
    return process72hrFollowUpCheck({
      triggerId: triggerId
    });
  } catch (err) {
    console.error("handle72hrFollowUp error:", err);
    return { success: false, message: err.toString() };
  }
}

function process72hrFollowUpCheck(options) {
  const props = PropertiesService.getScriptProperties();
  const triggerId = options.triggerId || null;
  let leadId = options.leadId || (triggerId ? props.getProperty(TRIGGER_LOOKUP_PREFIX + triggerId) : null);

  if (!leadId) {
    if (triggerId) deleteTriggerById(triggerId);
    return { success: false, message: "Lead not found for trigger." };
  }

  const storedJson = props.getProperty(FOLLOW_UP_PROPERTY_PREFIX + leadId);
  const stored = storedJson ? JSON.parse(storedJson) : {};

  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);

  let rowIndex = null;
  if (sheet) {
    rowIndex = findLeadRowByLeadId(sheet, leadId);
    if (!rowIndex && stored.email) {
      rowIndex = findLeadRowByEmail(sheet, stored.email);
    }
  }

  const bookedValue = rowIndex && sheet
    ? sheet.getRange(rowIndex, COL.BOOKED_CONSULTATION).getValue()
    : "";
  const isBookedConsultation = isTruthy(bookedValue);

  if (isBookedConsultation) {
    appendCloseAudit({
      timestamp: new Date(),
      action: "72hr-check",
      leadId: leadId,
      email: stored.email || "",
      name: stored.name || "",
      leadType: stored.leadType || "Cold",
      scorePercent: stored.scorePercent || "",
      promoCode: stored.promoCode || "",
      pageUrl: stored.sourceUrl || "",
      referrer: "",
      triggerId: triggerId || stored.triggerId || "",
      details: "Consultation already booked"
    });
    cleanupFollowUpArtifacts(leadId, triggerId);
    return { success: true, leadId: leadId, skipped: true };
  }

  if (stored.email) {
    send72hrFollowUpEmail({
      email: stored.email,
      name: stored.name || "",
      promoCode: stored.promoCode || "",
      leadId: leadId,
      scorePercent: stored.scorePercent || 0
    });
  }

  if (rowIndex && sheet) {
    mark72hrFollowUpSent({ rowIndex: rowIndex });
  }

  appendCloseAudit({
    timestamp: new Date(),
    action: "72hr-followup-sent",
    leadId: leadId,
    email: stored.email || "",
    name: stored.name || "",
    leadType: stored.leadType || "Cold",
    scorePercent: stored.scorePercent || "",
    promoCode: stored.promoCode || "",
    pageUrl: stored.sourceUrl || "",
    referrer: "",
    triggerId: triggerId || stored.triggerId || "",
    details: "72-hour follow-up email sent"
  });

  cleanupFollowUpArtifacts(leadId, triggerId);

  return { success: true, leadId: leadId, followUpSent: true };
}

function cleanupFollowUpArtifacts(leadId, triggerId) {
  const props = PropertiesService.getScriptProperties();
  let resolvedTriggerId = triggerId;

  if (!resolvedTriggerId && leadId) {
    const storedJson = props.getProperty(FOLLOW_UP_PROPERTY_PREFIX + leadId);
    if (storedJson) {
      try {
        const stored = JSON.parse(storedJson);
        resolvedTriggerId = stored.triggerId;
      } catch (err) {
        resolvedTriggerId = null;
      }
    }
  }

  if (leadId) {
    props.deleteProperty(FOLLOW_UP_PROPERTY_PREFIX + leadId);
  }
  if (resolvedTriggerId) {
    props.deleteProperty(TRIGGER_LOOKUP_PREFIX + resolvedTriggerId);
    deleteTriggerById(resolvedTriggerId);
  }
}

function deleteTriggerById(triggerId) {
  if (!triggerId) return;
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getUniqueId && trigger.getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function findLeadRowByLeadId(sheet, leadId) {
  if (!sheet || !leadId) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW_INDEX) return null;

  const values = sheet.getRange(HEADER_ROW_INDEX + 1, COL.LEAD_ID, lastRow - HEADER_ROW_INDEX, 1).getValues();
  const target = leadId.toString().trim().toLowerCase();

  for (let i = 0; i < values.length; i++) {
    const current = (values[i][0] || "").toString().trim().toLowerCase();
    if (current === target) {
      return HEADER_ROW_INDEX + 1 + i;
    }
  }

  return null;
}

function findLeadRowByEmail(sheet, email) {
  if (!sheet || !email) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW_INDEX) return null;

  const values = sheet.getRange(HEADER_ROW_INDEX + 1, COL.EMAIL, lastRow - HEADER_ROW_INDEX, 1).getValues();
  const target = email.toString().trim().toLowerCase();

  for (let i = 0; i < values.length; i++) {
    const current = (values[i][0] || "").toString().trim().toLowerCase();
    if (current === target) {
      return HEADER_ROW_INDEX + 1 + i;
    }
  }

  return null;
}

function getReportStatusSheet() {
  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(REPORT_STATUS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(REPORT_STATUS_SHEET_NAME);
    sheet.getRange(1, 1, 1, REPORT_STATUS_HEADERS.length).setValues([REPORT_STATUS_HEADERS]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, REPORT_STATUS_HEADERS.length).setValues([REPORT_STATUS_HEADERS]);
  }

  return sheet;
}

function findStatusRowByLeadId(sheet, leadId) {
  if (!sheet || !leadId) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const target = leadId.toString().trim().toLowerCase();

  for (let i = 0; i < values.length; i++) {
    const current = (values[i][0] || "").toString().trim().toLowerCase();
    if (current === target) {
      return 2 + i;
    }
  }

  return null;
}

function updateStatus(leadId, stepNumber, statusText) {
  try {
    if (!leadId) return;
    const sheet = getReportStatusSheet();
    if (!sheet) return;

    const rowIndex = findStatusRowByLeadId(sheet, leadId);
    const payload = [
      leadId,
      Number(stepNumber) || 0,
      statusText || "",
      new Date()
    ];

    if (rowIndex) {
      sheet.getRange(rowIndex, 1, 1, payload.length).setValues([payload]);
    } else {
      sheet.appendRow(payload);
    }
  } catch (err) {
    console.warn("updateStatus error", err);
  }
}

function getStatusByLeadId(leadId) {
  try {
    const sheet = getReportStatusSheet();
    if (!sheet || !leadId) {
      return { step: 0, statusText: "" };
    }

    const rowIndex = findStatusRowByLeadId(sheet, leadId);
    if (!rowIndex) {
      return { step: 0, statusText: "" };
    }

    const values = sheet.getRange(rowIndex, 1, 1, REPORT_STATUS_HEADERS.length).getValues();
    const [, stepNumber, statusText] = values[0];
    return {
      step: Number(stepNumber) || 0,
      statusText: statusText || ""
    };
  } catch (err) {
    console.warn("getStatusByLeadId error", err);
    return { step: 0, statusText: "" };
  }
}

function sendCloseAcknowledgementEmail(options) {
  const recipient = (options.email || "").trim();
  if (!recipient) return null;

  const firstName = (options.name || "").split(" ")[0] || "there";
  const subject = "We've saved your Bio-Balance Report";
  const replyTo = getScriptProp(CLOSE_EMAIL_FROM_PROP);

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;color:#111827;line-height:1.6;">
      <h2 style="color:#0a466b;margin:0 0 12px 0;">Hi ${sanitize(firstName)},</h2>
      <p style="margin:0 0 12px 0;">We noticed you closed your Bio-Balance report. We've saved it for you and can review it together.</p>
      <p style="margin:0 0 12px 0;">If you want quick answers or a treatment plan, book a free consultation below. Your ${options.promoCode ? "promo code " + sanitize(options.promoCode) + " is saved" : "report"} is on file.</p>
      <div style="margin:20px 0;">
        <a href="https://www.optimantra.com/optimus/patient/patientaccess/servicesall?pid=dkdSRlJnSnpWWWZmZ2J4Q3ExUUhCQT09&lid=RjBaSTBqc2tkV0FKSUVTRm9rS1k4UT09" target="_blank" style="background:#0a466b;color:#fff;padding:12px 18px;border-radius:8px;text-decoration:none;font-weight:700;display:inline-block;">Book your free consultation</a>
      </div>
      <p style="margin:0 0 8px 0;">If you have questions, reply to this email and we will help.</p>
      <p style="margin:0;color:#6b7280;font-size:13px;">White Tank Wellness Care Team</p>
    </div>
  `;

  const mailOptions = {
    to: recipient,
    subject: subject,
    htmlBody: htmlBody
  };

  if (replyTo) {
    mailOptions.replyTo = replyTo;
  }

  MailApp.sendEmail(mailOptions);
  return true;
}

function send72hrFollowUpEmail(options) {
  const recipient = (options.email || "").trim();
  if (!recipient) return null;

  const firstName = (options.name || "").split(" ")[0] || "there";
  const replyTo = getScriptProp(FOLLOW_UP_EMAIL_FROM_PROP) || getScriptProp(CLOSE_EMAIL_FROM_PROP);
  const subject = "Quick follow-up about your Bio-Balance Report";

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;color:#111827;line-height:1.6;">
      <h2 style="color:#0a466b;margin:0 0 12px 0;">Hi ${sanitize(firstName)},</h2>
      <p style="margin:0 0 12px 0;">It has been a few days since you generated your Bio-Balance report. If you would like to go over the findings or start a plan, our clinicians are ready to help.</p>
      <p style="margin:0 0 12px 0;">${options.promoCode ? "Your promo code " + sanitize(options.promoCode) + " is still active." : "We saved your report and can walk through it together."}</p>
      <div style="margin:20px 0;">
        <a href="https://www.optimantra.com/optimus/patient/patientaccess/servicesall?pid=dkdSRlJnSnpWWWZmZ2J4Q3ExUUhCQT09&lid=RjBaSTBqc2tkV0FKSUVTRm9rS1k4UT09" target="_blank" style="background:#0a466b;color:#fff;padding:12px 18px;border-radius:8px;text-decoration:none;font-weight:700;display:inline-block;">Schedule your consultation</a>
      </div>
      <p style="margin:0 0 8px 0;">If you already booked, thank you - no further action needed.</p>
      <p style="margin:0;color:#6b7280;font-size:13px;">White Tank Wellness Care Team</p>
    </div>
  `;

  const mailOptions = {
    to: recipient,
    subject: subject,
    htmlBody: htmlBody
  };

  if (replyTo) {
    mailOptions.replyTo = replyTo;
  }

  MailApp.sendEmail(mailOptions);
  return true;
}

function loadFollowUpTemplate() {
  const fallback = {
    templateName: "Default Follow-Up",
    subject: "Follow-Up Reminder for {{LeadName}} - {{FollowUpDate}}",
    body: [
      "Hi {{LeadName}},",
      "",
      "Just a quick reminder about our follow-up on {{FollowUpDate}}.",
      "Call or text us at {{Phone}} if you need to make a change or have questions.",
      "",
      "Talk soon,",
      "White Tank Wellness"
    ].join("\n"),
    cc: "",
    bcc: ""
  };

  try {
    const ss = SpreadsheetApp.getActive();
    if (!ss) {
      throw new Error("No active spreadsheet context available.");
    }
    const sheet = ss.getSheetByName("Email Templates");
    if (!sheet) {
      throw new Error("Sheet 'Email Templates' not found.");
    }

    const rowValues = sheet.getRange(2, 1, 1, 5).getValues();
    if (!rowValues || !rowValues.length) {
      throw new Error("No template data on row 2.");
    }
    const row = rowValues[0];

    return {
      templateName: row[0] || fallback.templateName,
      subject: row[1] || fallback.subject,
      body: row[2] || fallback.body,
      cc: row[3] || "",
      bcc: row[4] || ""
    };
  } catch (err) {
    Logger.log("loadFollowUpTemplate fallback: " + err);
    return fallback;
  }
}

function applyReminderTemplate(template, context) {
  if (typeof template !== "string") {
    return "";
  }
  if (!context) {
    return template;
  }
  return template.replace(/{{\s*(\w+)\s*}}/g, function(match, key) {
    if (Object.prototype.hasOwnProperty.call(context, key)) {
      return context[key] || "";
    }
    return "";
  });
}

function sendFollowUpReminder(data) {
  const payload = data || {};
  const toAddress = payload.to || payload.email || "";
  const to = String(toAddress).trim();
  if (!to) {
    throw new Error("sendFollowUpReminder requires a recipient address.");
  }

  const cc = payload.cc ? String(payload.cc).trim() : "";
  const bcc = payload.bcc ? String(payload.bcc).trim() : "";

  const leadName = (payload.leadName || payload.name || "Your Client").toString().trim() || "Your Client";
  const followUpDate = (payload.followUpDate || payload.followUpDateRaw || payload.dueInput || payload.dueDate || "").toString().trim() || "(date not specified)";
  const phone = (payload.phone || payload.leadPhone || "").toString().trim();
  const tagEmail = (payload.leadEmail || payload.emailForTags || payload.contactEmail || payload.email || "").toString().trim();

  const defaultSubject = "Reminder: Your Upcoming Follow-Up Task";
  const defaultBody = [
    "Hello {{LeadName}},",
    "",
    "This is a reminder for your scheduled follow-up on {{FollowUpDate}}.",
    "",
    "Call or text us at {{Phone}} if you have any questions.",
    "",
    "White Tank Wellness Lead Center"
  ].join("\n");

  const subjectTemplate = payload.subject && payload.subject.toString().trim() ? payload.subject.toString() : defaultSubject;
  const bodyTemplate = payload.body && payload.body.toString().trim() ? payload.body.toString() : defaultBody;

  const context = {
    LeadName: leadName,
    FollowUpDate: followUpDate,
    Phone: phone,
    Email: tagEmail
  };

  const subject = applyReminderTemplate(subjectTemplate, context) || defaultSubject;
  const body = applyReminderTemplate(bodyTemplate, context) || defaultBody;
  const htmlBody = body.replace(/\n/g, "<br>");

  const mailOptions = {
    to: to,
    subject: subject,
    body: body,
    htmlBody: htmlBody
  };

  if (cc) {
    mailOptions.cc = cc;
  }
  if (bcc) {
    mailOptions.bcc = bcc;
  }

  const replyTo = getScriptProp(FOLLOW_UP_EMAIL_FROM_PROP) || getScriptProp(CLOSE_EMAIL_FROM_PROP);
  if (replyTo) {
    mailOptions.replyTo = replyTo;
  }

  const attachments = Array.isArray(payload.attachments) ? payload.attachments : [];
  const blobs = attachments.reduce(function(list, file) {
    if (!file || !file.content) {
      return list;
    }
    try {
      const bytes = Utilities.base64Decode(String(file.content));
      const blob = Utilities.newBlob(bytes, file.mimeType || "application/octet-stream", file.name || "attachment");
      list.push(blob);
    } catch (err) {
      Logger.log("Skipping attachment: " + err);
    }
    return list;
  }, []);

  if (blobs.length) {
    mailOptions.attachments = blobs;
  }

  MailApp.sendEmail(mailOptions);

  return { success: true };
}

function sendEditableFollowUpReminder(payload) {
  const data = payload || {};
  const to = (data.to || data.email || "").toString().trim();
  if (!to) {
    throw new Error("sendEditableFollowUpReminder requires a recipient email address.");
  }

  const cc = (data.cc || "").toString().trim();
  const bcc = (data.bcc || "").toString().trim();
  const subject = (data.subject || "Follow-Up Reminder").toString();
  let body = (typeof data.body === "string" && data.body.trim()) ? data.body : "";

  const tokens = {
    "{{LeadName}}": data.leadName || "",
    "{{FollowUpDate}}": data.followUpDate || data.followUpDateRaw || "",
    "{{Phone}}": data.leadPhone || data.phone || "",
    "{{Email}}": to
  };

  Object.keys(tokens).forEach(function(tag) {
    const value = tokens[tag] || "";
    if (body) {
      body = body.split(tag).join(value);
    }
  });

  const attachmentsInput = Array.isArray(data.attachments) ? data.attachments : [];
  const attachments = attachmentsInput.reduce(function(list, file) {
    if (!file) {
      return list;
    }
    const base64 = file.data || file.content || "";
    if (!base64) {
      return list;
    }
    try {
      const decoded = Utilities.base64Decode(base64);
      const blob = Utilities.newBlob(decoded, file.type || "application/octet-stream", file.name || "attachment");
      list.push(blob);
    } catch (err) {
      Logger.log("Failed to decode attachment: " + err);
    }
    return list;
  }, []);

  const mailOptions = {
    to: to,
    subject: subject,
    htmlBody: body || undefined,
    body: body ? body.replace(/<[^>]+>/g, "") : undefined
  };

  if (cc) {
    mailOptions.cc = cc;
  }
  if (bcc) {
    mailOptions.bcc = bcc;
  }
  if (attachments.length) {
    mailOptions.attachments = attachments;
  }

  MailApp.sendEmail(mailOptions);
  return { success: true };
}

function sendIntakeReportEmail(data) {
  const leadId = data.leadId;
  if (!leadId) {
    throw new Error("sendIntakeReportEmail called without a leadId.");
  }

  const rowIndex = data.rowIndex;
  if (!rowIndex) {
    throw new Error("sendIntakeReportEmail called without a rowIndex.");
  }

  // Generate report HTML using your existing pipeline
  const result = generateReportForLead({
    rowIndex: rowIndex,
    sheetName: SUBMISSIONS_SHEET_NAME,
    sendEmail: false
  });

  let reportHtml = result.reportHtml || result;

  //
  // 1) EMAIL REPORT TO PATIENT
  //
  if (data.email && data.email.trim() !== "") {
    MailApp.sendEmail({
      to: data.email,
      subject: "Your Bio-Balance Method™ Report",
      htmlBody: reportHtml
    });
  }

  //
  // 2) SAVE REPORT TO GOOGLE DRIVE FOLDER
  //
  const folderName = "Patient Hormone Assessment Reports";
  let folder;

  // Find folder or create if missing
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  // Create filename
  const safeName = (data.name || "Patient") + " - " + leadId + ".pdf";

  // Convert HTML → PDF blob
  const dateString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy");
  const patientName = data.name || "Patient";
  const score = data.bioBalanceScore || 0;

  const pdfHeader = `
  <div style="width:100%; padding:20px; border-bottom:2px solid #0a466b20;
              font-family:Arial; margin-bottom:20px;">
      <div style="font-size:22px; font-weight:700; color:#0a466b;">
        Bio-Balance Method™ Assessment Report
      </div>
      <div style="margin-top:6px; color:#333; font-size:14px;">
        <strong>Patient:</strong> ${patientName}<br>
        <strong>Date:</strong> ${dateString}<br>
        <strong>Bio-Balance Score:</strong> ${score}/100
      </div>
  </div>
`;

  reportHtml = pdfHeader + reportHtml;
  const blob = Utilities.newBlob(reportHtml, "text/html", "report.html");
  const pdf = blob.getAs("application/pdf");

  // Save
  const file = folder.createFile(pdf).setName(safeName);

  // Insert PDF URL into the sheet
  const ss = SpreadsheetApp.openById(SUBMISSIONS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Submissions");

  // Ensure Report PDF Link column exists (make it column 57 if free)
  const pdfColIndex = 57;
  ensureColumnExists(sheet, pdfColIndex);

  // Save the PDF URL to the sheet
  if (rowIndex) {
    const fileUrl = file.getUrl();
    sheet.getRange(rowIndex, pdfColIndex).setValue(fileUrl);
  }
}

function saveInternalIntake(data) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Submissions');

  const row = sheet.getLastRow() + 1;

  // Core identifiers
  const timestamp = new Date();
  const leadId = 'INT-' + timestamp.getTime();

  // Symptom-based scoring
  const symptomCount = Array.isArray(data.symptoms) ? data.symptoms.length : 0;
  const maxSymptoms = 10; // number of symptom checkboxes in the sidebar
  const bioBalanceScore = maxSymptoms > 0
    ? Math.round((symptomCount / maxSymptoms) * 100)
    : 0;

  // Simple bucket based on score
  let coldBucket = '';
  if (bioBalanceScore >= 80) {
    coldBucket = 'Hot - Strong Symptoms';
  } else if (bioBalanceScore >= 40) {
    coldBucket = 'Warm - Moderate Symptoms';
  } else {
    coldBucket = 'Mild - Monitor';
  }

  const leadType = 'Cold Phone - Internal Intake';
  const promoCode = (data.promoCode || '').toString().trim().toUpperCase();
  const promoDescription = promoCode ? autoPromoDescription(promoCode) : '';

  data.promoCode = promoCode;

  // Write to "Submissions" in the exact column order of the header row
  sheet.getRange(row, 1).setValue(timestamp);                           // Timestamp
  sheet.getRange(row, 2).setValue(leadId);                              // Lead ID
  sheet.getRange(row, 3).setValue('Internal Call');                     // Source
  sheet.getRange(row, 4).setValue('');                                  // UTM Campaign
  sheet.getRange(row, 5).setValue('');                                  // UTM Ad Set
  sheet.getRange(row, 6).setValue('');                                  // UTM Keyword
  sheet.getRange(row, 7).setValue('Internal Intake Sidebar');           // UTM Landing Page
  sheet.getRange(row, 8).setValue(data.gender || '');                   // Gender
  sheet.getRange(row, 9).setValue(data.name || '');                     // Name
  sheet.getRange(row,10).setValue(data.email || '');                    // Email
  sheet.getRange(row,11).setValue(data.phone || '');                    // Phone
  sheet.getRange(row,12).setValue(data.ageRange || '');                 // Age Range

  sheet.getRange(row,13).setValue(data.sleepQuality || '');             // Sleep Quality
  sheet.getRange(row,14).setValue(data.energyLevels || '');             // Energy Levels
  sheet.getRange(row,15).setValue(data.mentalClarity || '');            // Mental Clarity
  sheet.getRange(row,16).setValue(data.stressTolerance || '');          // Stress Tolerance
  sheet.getRange(row,17).setValue(data.physicalActivity || '');         // Physical Activity
  sheet.getRange(row,18).setValue(data.bodyChanges || '');              // Body Changes
  sheet.getRange(row,19).setValue(data.libido || '');                   // Libido
  sheet.getRange(row,20).setValue(data.temperatureSensitivity || '');   // Temperature Sensitivity (F)
  sheet.getRange(row,21).setValue(data.cycleStatus || '');              // Cycle Status (F)
  sheet.getRange(row,22).setValue(data.pmsMoodShifts || '');            // PMS / Mood Shifts (F)
  sheet.getRange(row,23).setValue(data.bloating || '');                 // Bloating / Fluid Retention (F)
  sheet.getRange(row,24).setValue(data.strengthRecovery || '');         // Strength & Recovery (M)
  sheet.getRange(row,25).setValue(data.morningDrive || '');             // Morning Drive / Motivation (M)
  sheet.getRange(row,26).setValue(data.muscleMass || '');               // Muscle Mass / Strength Changes (M)
  sheet.getRange(row,27).setValue(data.sexualPerformance || '');        // Sexual Performance / Libido (M)

  sheet.getRange(row,28).setValue(bioBalanceScore);                     // Bio-Balance Score (%)
  sheet.getRange(row,29).setValue('');                                  // Lead Score (Sales)
  sheet.getRange(row,30).setValue(promoCode);                           // Promo Code
  sheet.getRange(row,31).setValue(promoDescription);                    // Promo Description
  sheet.getRange(row,32).setValue('');                                  // Promo Redeemed?
  sheet.getRange(row,33).setValue('');                                  // B12 Offered?
  sheet.getRange(row,34).setValue('');                                  // B12 Redeemed?

  sheet.getRange(row,35).setValue('New - Internal Call');               // Lead Status
  sheet.getRange(row,36).setValue('');                                  // Assigned Staff
  sheet.getRange(row,37).setValue('');                                  // Follow-Up Priority
  sheet.getRange(row,38).setValue('');                                  // Last Outreach Date
  sheet.getRange(row,39).setValue('');                                  // Next Follow-Up Date
  sheet.getRange(row,40).setValue(0);                                   // Contact Attempts
  sheet.getRange(row,41).setValue('');                                  // Last Attempt Method
  sheet.getRange(row,42).setValue('');                                  // Last Attempt Result
  sheet.getRange(row,43).setValue('No');                                // Opted Out?

  sheet.getRange(row,44).setValue(data.notes || '');                    // Notes
  sheet.getRange(row,45).setValue('');                                  // Booked Consultation?
  sheet.getRange(row,46).setValue('');                                  // Consultation Date
  sheet.getRange(row,47).setValue('');                                  // Joined Membership?
  sheet.getRange(row,48).setValue('');                                  // Membership Type
  sheet.getRange(row,49).setValue('');                                  // Membership Start Date

  sheet.getRange(row,50).setValue(coldBucket);                          // Cold Lead Retargeting Bucket
  sheet.getRange(row,51).setValue(
    Array.isArray(data.tags) ? data.tags.join(', ') : ''
  );                                                                    // Tags

  // Determine if a report should be generated for this save action
  const shouldGenerateReport = !(
    Object.prototype.hasOwnProperty.call(data, 'generateReport') &&
    (data.generateReport === false || data.generateReport === 'false')
  );

  if (Object.prototype.hasOwnProperty.call(data, 'generateReport')) {
    delete data.generateReport;
  }

  const rawAnswers = { ...data, leadId: leadId };
  sheet.getRange(row,52).setValue(JSON.stringify(rawAnswers));         // Raw Answers (JSON)
  sheet.getRange(row,53).setValue(leadType);                            // Lead Type
  sheet.getRange(row,54).setValue('');                                  // Follow-Up Sent?
  sheet.getRange(row,55).setValue('');                                  // 72hr Follow-Up Timestamp
  sheet.getRange(row,56).setValue('');                                  // Closed From

  // pass enriched data and row index into the report generator when requested
  if (shouldGenerateReport) {
    sendIntakeReportEmail({
      ...data,
      leadId: leadId,
      bioBalanceScore: bioBalanceScore,
      rowIndex: row
    });
  }

  return {
    leadId: leadId,
    reportGenerated: shouldGenerateReport
  };
}
