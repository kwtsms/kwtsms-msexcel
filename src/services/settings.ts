// src/services/settings.ts
// Wrapper for Office.context.document.settings (per-document, persists when saved).
// 2026-11-26 13:00 +03: except the login, which is kept in localStorage (see saveCredentials), #1758
// All writes call saveAsync() to persist to document.
// Related: src/models/types.ts

/* global window */ // 2026-11-26 13:05 +03: browser global for the localStorage login, lint no-undef, #1758
import { KwtSmsCredentials, CachedData } from "../models/types";

const KEYS = {
  USERNAME: "kwtsms_username",
  PASSWORD: "kwtsms_password",
  BALANCE: "kwtsms_balance",
  SENDER_IDS: "kwtsms_senderids",
  COVERAGE: "kwtsms_coverage",
  TEST_MODE: "kwtsms_testmode",
  DEFAULT_COUNTRY: "kwtsms_default_country",
  SELECTED_SENDER: "kwtsms_selected_sender",
};

function get(key: string): any {
  return Office.context.document.settings.get(key);
}

function set(key: string, value: any): void {
  Office.context.document.settings.set(key, value);
}

function remove(key: string): void {
  Office.context.document.settings.remove(key);
}

function saveAsync(): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error?.message || "Failed to save settings"));
      }
    });
  });
}

// 2026-11-26 13:00 +03: credentials live in the task pane's localStorage (per user, per device), never in
// document settings: those are saved inside the .xlsx, so anyone the sheet was shared with could read the
// API password. Security review #1758 c17608 (HIGH), Microsoft guidance; #968.
function localGet(key: string): string | null {
  try {
    return window.localStorage.getItem(key);
  } catch {
    return null;
  }
}

function localSet(key: string, value: string): void {
  window.localStorage.setItem(key, value);
}

function localRemove(key: string): void {
  try {
    window.localStorage.removeItem(key);
  } catch {
    // storage blocked: nothing stored, nothing to remove
  }
}

// 2026-11-26 13:10 +03: files saved by 1.0.20 and older carry the login. Remove it, never copy it: the file may
// come from someone else, and copying would send on their credits from this device (Onyx #1758 c17630).
// Returns true when a login was removed, so the task pane can ask the user to log in again.
export function removeFileLogin(): boolean {
  if (!get(KEYS.USERNAME) && !get(KEYS.PASSWORD)) return false;
  remove(KEYS.USERNAME);
  remove(KEYS.PASSWORD);
  saveAsync().catch(() => undefined);
  return true;
}

export async function saveCredentials(username: string, password: string): Promise<void> {
  localSet(KEYS.USERNAME, username);
  localSet(KEYS.PASSWORD, password);
}

export function getCredentials(): KwtSmsCredentials | null {
  const username = localGet(KEYS.USERNAME);
  const password = localGet(KEYS.PASSWORD);
  if (!username || !password) return null;
  return { username, password };
}

export async function clearCredentials(): Promise<void> {
  localRemove(KEYS.USERNAME);
  localRemove(KEYS.PASSWORD);
  remove(KEYS.USERNAME);
  remove(KEYS.PASSWORD);
  await saveAsync();
}

export async function saveCachedData(balance: number, senderIds: string[], coverage: string[]): Promise<void> {
  set(KEYS.BALANCE, balance);
  set(KEYS.SENDER_IDS, JSON.stringify(senderIds));
  set(KEYS.COVERAGE, JSON.stringify(coverage));
  await saveAsync();
}

export function getCachedData(): CachedData | null {
  const balance = get(KEYS.BALANCE);
  const senderIdsRaw = get(KEYS.SENDER_IDS);
  const coverageRaw = get(KEYS.COVERAGE);
  if (balance === null || balance === undefined) return null;
  return {
    balance: Number(balance),
    senderIds: senderIdsRaw ? JSON.parse(senderIdsRaw) : [],
    coverage: coverageRaw ? JSON.parse(coverageRaw) : [],
  };
}

export async function updateBalance(balance: number): Promise<void> {
  set(KEYS.BALANCE, balance);
  await saveAsync();
}

export async function clearCachedData(): Promise<void> {
  remove(KEYS.BALANCE);
  remove(KEYS.SENDER_IDS);
  remove(KEYS.COVERAGE);
  await saveAsync();
}

export async function setTestMode(enabled: boolean): Promise<void> {
  set(KEYS.TEST_MODE, enabled);
  await saveAsync();
}

export function getTestMode(): boolean {
  const val = get(KEYS.TEST_MODE);
  return val === null || val === undefined ? true : Boolean(val);
}

export async function setDefaultCountryCode(code: string): Promise<void> {
  set(KEYS.DEFAULT_COUNTRY, code);
  await saveAsync();
}

export function getDefaultCountryCode(): string {
  return get(KEYS.DEFAULT_COUNTRY) || "965";
}

export async function setSelectedSenderId(senderId: string): Promise<void> {
  set(KEYS.SELECTED_SENDER, senderId);
  await saveAsync();
}

export function getSelectedSenderId(): string {
  return get(KEYS.SELECTED_SENDER) || "KWT-SMS";
}
