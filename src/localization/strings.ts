// src/localization/strings.ts
// All UI strings keyed by locale. English for v1, Arabic placeholder for v1.1.
// Related: src/taskpane/taskpane.ts

export interface LocaleStrings {
  loginTitle: string;
  usernamePlaceholder: string;
  passwordPlaceholder: string;
  loginButton: string;
  loginError: string;
  // 2026-11-26 12:08 +03: network failure text, #968
  networkError: string;
  // 2026-11-26 13:10 +03: old workbook login removed, #1758
  fileLoginRemoved: string;
  loggingIn: string;
  balanceLabel: string;
  senderIdLabel: string;
  testModeLabel: string;
  logoutButton: string;
  phoneColumnLabel: string;
  messageColumnLabel: string;
  refreshColumnsButton: string;
  countryCodeLabel: string;
  messageTemplateLabel: string;
  messageTemplatePlaceholder: string;
  sendButton: string;
  sending: string;
  previewValid: string;
  previewDuplicates: string;
  previewSkipped: string;
  resultSent: string;
  resultFailed: string;
  resultSkipped: string;
  resultPointsUsed: string;
  // 2026-11-26 12:53 +03: account Test mode note, #968
  accountTestMode: string;
  noPhoneColumn: string;
  noValidNumbers: string;
  insufficientBalance: string;
  emptyMessage: string;
  supportLink: string;
  version: string;
}

const en: LocaleStrings = {
  loginTitle: "Login to kwtSMS API",
  usernamePlaceholder: "API Username",
  passwordPlaceholder: "API Password",
  loginButton: "Login",
  loginError: "Invalid credentials or service unavailable",
  // 2026-11-26 12:08 +03: shown when the request cannot reach kwtSMS at all, #968
  networkError:
    "Could not reach kwtSMS. Check your internet connection and try again. If it keeps failing, write to support@kwtsms.com.",
  // 2026-11-26 13:10 +03: shown once when a login saved in this workbook by an older version was removed, #1758
  fileLoginRemoved:
    "Please log in again. Your kwtSMS login is no longer kept in the workbook, so sharing a file never shares it.",
  loggingIn: "Logging in...",
  balanceLabel: "Balance",
  senderIdLabel: "Sender ID",
  testModeLabel: "Test Mode",
  logoutButton: "Logout",
  phoneColumnLabel: "Phone column",
  messageColumnLabel: "Message column (optional)",
  refreshColumnsButton: "Refresh columns",
  countryCodeLabel: "Country code",
  messageTemplateLabel: "Message template",
  messageTemplatePlaceholder: "Type your message here (used for all recipients if no message column selected)",
  sendButton: "Send SMS",
  sending: "Sending...",
  previewValid: "{count} valid numbers",
  previewDuplicates: "{count} duplicates removed",
  previewSkipped: "{count} skipped (no coverage)",
  resultSent: "Sent: {count}",
  resultFailed: "Failed: {count}",
  resultSkipped: "Skipped: {count}",
  resultPointsUsed: "Points used: {count}",
  // 2026-11-26 12:53 +03: shown when the send answer has "test": true but the add-in sent test=0, #968
  accountTestMode:
    "Your kwtSMS account is in Test mode: the SMS was queued as a test and not delivered. Turn Test mode off in your kwtSMS account to deliver.",
  noPhoneColumn: "Select a phone column",
  noValidNumbers: "No valid phone numbers found",
  insufficientBalance: "Insufficient balance: {needed} messages to send, {available} points available",
  emptyMessage: "Message is empty after cleaning",
  supportLink: "Support",
  version: "v{version}",
};

const LOCALES: Record<string, LocaleStrings> = { en };

export function getString(key: keyof LocaleStrings, locale: string = "en"): string {
  const strings = LOCALES[locale] || LOCALES.en;
  return strings[key];
}

export function formatString(template: string, values: Record<string, string | number>): string {
  let result = template;
  for (const [key, value] of Object.entries(values)) {
    result = result.replace("{" + key + "}", String(value));
  }
  return result;
}
