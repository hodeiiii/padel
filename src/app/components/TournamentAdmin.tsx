"use client";

import {
  AlertTriangle,
  CalendarClock,
  CheckCircle2,
  Copy,
  ExternalLink,
  FileSpreadsheet,
  LockKeyhole,
  Medal,
  Plus,
  RotateCcw,
  Share2,
  Trash2,
  Trophy,
  Upload,
  Users,
} from "lucide-react";
import {
  useEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
  type ComponentType,
  type CSSProperties,
  type ReactNode,
} from "react";
import * as XLSX from "xlsx";

type SheetCell = string | number | boolean | null | undefined;

type Pair = {
  id: string;
  seed: number;
  playerOne: string;
  playerTwo: string;
  restriction: string;
  comment: string;
  category: string;
};

type CategoryData = {
  id: string;
  name: string;
  pairs: Pair[];
  warnings: string[];
};

type Team = {
  id: string;
  seed?: number;
  name: string;
  restriction?: string;
  isPlaceholder?: boolean;
};

type Match = {
  id: string;
  label: string;
  sideA: Team | null;
  sideB: Team | null;
  winner: Team | null;
  loser: Team | null;
};

type DayKey = "jueves" | "viernes" | "sabado";

type ScheduleSlot = {
  day: DayKey;
  dayLabel: string;
  time: string;
  minutes: number;
  court: number;
  slotIndex: number;
};

type ScheduleTimeGroup = {
  groupIndex: number;
  slots: ScheduleSlot[];
};

type ScheduleAssignment = ScheduleSlot & {
  conflict?: boolean;
  manual?: boolean;
};

type ScheduleResult = {
  assignments: Record<string, ScheduleAssignment>;
  saturdayCount: number;
  total: number;
  conflicts: number;
};

type Round = {
  name: string;
  matches: Match[];
};

type Draw = {
  title: string;
  accent: "main" | "consolation";
  rounds: Round[];
};

type CategoryDrawSet = {
  categoryId: string;
  mainDraw: Draw;
  consolationDraw: Draw;
};

type SelectionMap = Record<string, string>;
type CategorySelectionMap = Record<string, SelectionMap>;
type AdminTab = "draws" | "summary";

type ManualScheduleOverride = {
  day: DayKey;
  time: string;
  court: number;
};

type ManualScheduleMap = Record<string, ManualScheduleOverride>;

type TournamentStateSnapshot = {
  categories: CategoryData[];
  selectedCategoryId: string;
  mainSelectionsByCategory: CategorySelectionMap;
  consolationSelectionsByCategory: CategorySelectionMap;
  manualScheduleOverrides: ManualScheduleMap;
};

type TournamentDraft = TournamentStateSnapshot & {
  publishedSlug?: string;
  savedAt: string;
};

type PublishedTournament = TournamentStateSnapshot & {
  slug: string;
  title: string;
  publishedAt: string;
  updatedAt: string;
};

type InitialAdminState = TournamentStateSnapshot & {
  publishedSlug: string;
};

const previewLimit = 24;
const courtCount = 8;
const bracketMatchHeight = 210;
const bracketBaseGap = 16;
const draftStorageKey = "padel-admin-draft-v1";
const activePublicationStorageKey = "padel-active-publication-v1";
const adminAuthStorageKey = "padel-admin-auth-v1";
const adminPassword = "landerlander";
const dayOrder: Record<DayKey, number> = {
  jueves: 0,
  viernes: 1,
  sabado: 2,
};
const emptyMainDraw: Draw = {
  title: "Cuadro principal",
  accent: "main",
  rounds: [],
};
const emptyConsolationDraw: Draw = {
  title: "Consolacion",
  accent: "consolation",
  rounds: [],
};
const emptyInitialAdminState: InitialAdminState = {
  categories: [],
  selectedCategoryId: "",
  mainSelectionsByCategory: {},
  consolationSelectionsByCategory: {},
  manualScheduleOverrides: {},
  publishedSlug: "",
};
const publishedTournamentCache = new Map<
  string,
  {
    raw: string | null;
    value: PublishedTournament | null;
  }
>();

type RoundLayoutStyle = CSSProperties & {
  "--round-gap": string;
  "--round-padding": string;
};

function bracketRoundLayout(roundIndex: number) {
  const stride = bracketMatchHeight + bracketBaseGap;

  return {
    gap: `${2 ** roundIndex * stride - bracketMatchHeight}px`,
    paddingTop: `${((2 ** roundIndex - 1) * stride) / 2}px`,
  };
}

function bracketRoundStyle(roundIndex: number): RoundLayoutStyle {
  const layout = bracketRoundLayout(roundIndex);

  return {
    "--round-gap": layout.gap,
    "--round-padding": layout.paddingTop,
  };
}

function domSlug(value: string) {
  return normalizeText(value).replace(/[^a-z0-9]+/g, "-");
}

function publishedTournamentStorageKey(slug: string) {
  return `padel-public-tournament-${slug}`;
}

function firstSurname(fullName: string) {
  const tokens = fullName.trim().split(/\s+/).filter(Boolean);

  if (tokens.length <= 1) return tokens[0] || "";

  const firstSurnameIndex = 1;
  const firstSurnameToken = tokens[firstSurnameIndex];
  const particles = new Set(["de", "del", "da", "dos", "das", "la", "las"]);

  if (
    particles.has(normalizeText(firstSurnameToken)) &&
    tokens[firstSurnameIndex + 1]
  ) {
    return `${firstSurnameToken} ${tokens[firstSurnameIndex + 1]}`;
  }

  return firstSurnameToken;
}

function cleanCell(value: SheetCell) {
  return String(value ?? "").trim();
}

function properCaseName(value: string) {
  return value
    .trim()
    .replace(/\s+/g, " ")
    .split(" ")
    .map((word) =>
      word
        .split("-")
        .map((segment) => {
          const lower = segment.toLocaleLowerCase("es-ES");

          return lower
            ? lower.charAt(0).toLocaleUpperCase("es-ES") + lower.slice(1)
            : "";
        })
        .join("-"),
    )
    .join(" ");
}

function normalizeText(value: string) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function nextPowerOfTwo(value: number) {
  let size = 1;

  while (size < value) size *= 2;

  return Math.max(size, 2);
}

function makeCategoryId(name: string, index: number) {
  return `${normalizeText(name).replace(/[^a-z0-9]+/g, "-") || "categoria"}-${index}`;
}

function makeTournamentSlug() {
  return `torneo-padel-${Date.now().toString(36)}`;
}

function tournamentTitle(categories: CategoryData[]) {
  if (categories.length === 1) return categories[0].name;

  return categories.length
    ? `Torneo de padel (${categories.length} categorias)`
    : "Torneo de padel";
}

function readStorageJson<T>(key: string) {
  if (typeof window === "undefined") return null;

  try {
    const raw = window.localStorage.getItem(key);

    return raw ? (JSON.parse(raw) as T) : null;
  } catch {
    return null;
  }
}

function writeStorageJson(key: string, value: unknown) {
  if (typeof window === "undefined") return;

  window.localStorage.setItem(key, JSON.stringify(value));
}

function removeStorageItem(key: string) {
  if (typeof window === "undefined") return;

  window.localStorage.removeItem(key);
}

function withoutCategoryManualSchedules(
  overrides: ManualScheduleMap,
  categoryId: string,
) {
  return Object.fromEntries(
    Object.entries(overrides).filter(
      ([key]) => !key.startsWith(`${categoryId}::`),
    ),
  ) as ManualScheduleMap;
}

function buildTournamentSnapshot({
  categories,
  consolationSelectionsByCategory,
  mainSelectionsByCategory,
  manualScheduleOverrides,
  selectedCategoryId,
}: TournamentStateSnapshot): TournamentStateSnapshot {
  return {
    categories,
    consolationSelectionsByCategory,
    mainSelectionsByCategory,
    manualScheduleOverrides,
    selectedCategoryId,
  };
}

function savePublishedTournament(
  slug: string,
  snapshot: TournamentStateSnapshot,
) {
  const key = publishedTournamentStorageKey(slug);
  const existing = readStorageJson<PublishedTournament>(key);
  const now = new Date().toISOString();
  const tournament: PublishedTournament = {
    ...snapshot,
    slug,
    title: tournamentTitle(snapshot.categories),
    publishedAt: existing?.publishedAt ?? now,
    updatedAt: now,
  };

  writeStorageJson(key, tournament);
  writeStorageJson(activePublicationStorageKey, slug);
  publishedTournamentCache.set(slug, {
    raw: JSON.stringify(tournament),
    value: tournament,
  });

  return tournament;
}

function readPublishedTournament(slug: string) {
  if (typeof window === "undefined") return null;

  const key = publishedTournamentStorageKey(slug);
  const raw = window.localStorage.getItem(key);
  const cached = publishedTournamentCache.get(slug);

  if (cached?.raw === raw) return cached.value;

  try {
    const value = raw ? (JSON.parse(raw) as PublishedTournament) : null;

    publishedTournamentCache.set(slug, {
      raw,
      value,
    });

    return value;
  } catch {
    publishedTournamentCache.set(slug, {
      raw,
      value: null,
    });

    return null;
  }
}

function readInitialAdminState(): InitialAdminState {
  const draft = readStorageJson<TournamentDraft>(draftStorageKey);
  const publishedSlug =
    draft?.publishedSlug ||
    readStorageJson<string>(activePublicationStorageKey) ||
    "";

  return draft?.categories.length
    ? {
        categories: draft.categories,
        consolationSelectionsByCategory:
          draft.consolationSelectionsByCategory ?? {},
        mainSelectionsByCategory: draft.mainSelectionsByCategory ?? {},
        manualScheduleOverrides: draft.manualScheduleOverrides ?? {},
        publishedSlug,
        selectedCategoryId: draft.selectedCategoryId || draft.categories[0].id,
      }
    : {
        ...emptyInitialAdminState,
        publishedSlug,
      };
}

function useClientReady() {
  return useSyncExternalStore(
    () => () => undefined,
    () => true,
    () => false,
  );
}

function subscribeStorageKey(key: string, onStoreChange: () => void) {
  if (typeof window === "undefined") return () => undefined;

  function handleStorage(event: StorageEvent) {
    if (event.key === key) onStoreChange();
  }

  window.addEventListener("storage", handleStorage);
  const refreshTimer = window.setInterval(onStoreChange, 2000);

  return () => {
    window.clearInterval(refreshTimer);
    window.removeEventListener("storage", handleStorage);
  };
}

function findCategoryName(cells: string[]) {
  return cells.find((cell) => normalizeText(cell).includes("categoria")) || "";
}

function parseCategories(rows: SheetCell[][]) {
  const categories: CategoryData[] = [];
  let current: CategoryData | null = null;

  const openCategory = (name: string) => {
    current = {
      id: makeCategoryId(name, categories.length + 1),
      name,
      pairs: [],
      warnings: [],
    };
    categories.push(current);
  };

  rows.forEach((row, rowIndex) => {
    const cells = row.map(cleanCell);
    const categoryName = findCategoryName(cells);

    if (categoryName) {
      openCategory(categoryName);
      return;
    }

    const playerOne = properCaseName(cells[1] || "");
    const playerTwo = properCaseName(cells[2] || "");

    if (!playerOne && !playerTwo) return;

    if (!current) openCategory("Categoria unica");

    const activeCategory = current;

    if (!activeCategory) return;

    const seed = Number.parseInt(cells[0], 10);
    const pair: Pair = {
      id: `pair-${rowIndex}-${activeCategory.pairs.length + 1}`,
      seed: Number.isFinite(seed) ? seed : activeCategory.pairs.length + 1,
      playerOne,
      playerTwo,
      restriction: cells[3] || "",
      comment: cells[4] || "",
      category: activeCategory.name,
    };

    if (!pair.playerOne || !pair.playerTwo) {
      activeCategory.warnings.push(
        `Fila ${rowIndex + 1}: pareja incompleta en columnas B/C.`,
      );
    }

    activeCategory.pairs.push(pair);
  });

  return categories.filter((category) => category.pairs.length > 0);
}

function pairName(pair: Pair) {
  const playerOne = firstSurname(pair.playerOne) || pair.playerOne || "Jugador 1";
  const playerTwo = firstSurname(pair.playerTwo) || pair.playerTwo || "Jugador 2";

  return `${playerOne} / ${playerTwo}`;
}

function pairToTeam(pair: Pair): Team {
  return {
    id: pair.id,
    seed: pair.seed,
    name: pairName(pair),
    restriction: pair.restriction,
  };
}

function isPlayable(team: Team | null) {
  return Boolean(team && !team.isPlaceholder);
}

function roundName(roundIndex: number, totalRounds: number) {
  const remaining = totalRounds - roundIndex;

  if (remaining === 1) return "Final";
  if (remaining === 2) return "Semifinales";
  if (remaining === 3) return "Cuartos";
  if (remaining === 4) return "Octavos";

  return `Ronda ${roundIndex + 1}`;
}

function buildDraw(
  title: string,
  accent: Draw["accent"],
  teams: Team[],
  selections: SelectionMap,
  prefix: string,
) {
  if (teams.length < 2) {
    return { title, accent, rounds: [] };
  }

  const size = nextPowerOfTwo(teams.length);
  const totalRounds = Math.log2(size);
  const slots = Array<Team | null>(size).fill(null);

  teams.forEach((team, index) => {
    slots[index] = team;
  });

  let inputs = slots;
  let matchNumber = 1;
  const rounds: Round[] = [];

  for (let roundIndex = 0; roundIndex < totalRounds; roundIndex += 1) {
    const matches: Match[] = [];
    const nextInputs: (Team | null)[] = [];
    const matchesInRound = inputs.length / 2;

    for (let index = 0; index < matchesInRound; index += 1) {
      const sideA = inputs[index * 2] || null;
      const sideB = inputs[index * 2 + 1] || null;
      const id = `${prefix}${matchNumber}`;
      const selectedId = selections[id];
      const selectedSide =
        sideA?.id === selectedId ? sideA : sideB?.id === selectedId ? sideB : null;
      const automaticWinner =
        isPlayable(sideA) && !sideB
          ? sideA
          : isPlayable(sideB) && !sideA
            ? sideB
            : null;
      const winner = selectedSide || automaticWinner;
      const loser =
        winner && isPlayable(sideA) && isPlayable(sideB)
          ? winner.id === sideA?.id
            ? sideB
            : sideA
          : null;
      const match: Match = {
        id,
        label: `${roundName(roundIndex, totalRounds)} ${index + 1}`,
        sideA,
        sideB,
        winner,
        loser,
      };

      matches.push(match);
      nextInputs.push(
        winner || {
          id: `winner-${id}`,
          name: `Ganador ${id}`,
          isPlaceholder: true,
        },
      );
      matchNumber += 1;
    }

    rounds.push({
      name: roundName(roundIndex, totalRounds),
      matches,
    });
    inputs = nextInputs;
  }

  return { title, accent, rounds };
}

function consolationTeamsFrom(mainDraw: Draw) {
  const firstRound = mainDraw.rounds[0]?.matches ?? [];

  return firstRound
    .filter((match) => isPlayable(match.sideA) && isPlayable(match.sideB))
    .map(
      (match) =>
        match.loser || {
          id: `loser-${match.id}`,
          name: `Perdedor ${match.id}`,
          isPlaceholder: true,
        },
    );
}

function buildDrawSets(
  categories: CategoryData[],
  mainSelectionsByCategory: CategorySelectionMap,
  consolationSelectionsByCategory: CategorySelectionMap,
) {
  return categories.map((category) => {
    const mainDraw = buildDraw(
      "Cuadro principal",
      "main",
      category.pairs.map(pairToTeam),
      mainSelectionsByCategory[category.id] ?? {},
      "P",
    );
    const consolationDraw = buildDraw(
      "Consolacion",
      "consolation",
      consolationTeamsFrom(mainDraw),
      consolationSelectionsByCategory[category.id] ?? {},
      "C",
    );

    return {
      categoryId: category.id,
      mainDraw,
      consolationDraw,
    };
  });
}

function buildCsvTemplate() {
  return [
    ["", "1a CATEGORIA", "", "", ""],
    [1, "Jugador A", "Jugador B", "jueves y viernes + 18:00", "comentario"],
    [2, "Jugador C", "Jugador D", "viernes + 17:00", ""],
    ["", "2a CATEGORIA", "", "", ""],
    [1, "Jugador E", "Jugador F", "jueves NO", ""],
  ];
}

function escapeCsv(value: string | number) {
  const text = String(value);

  if (/[",\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;

  return text;
}

function downloadTemplate() {
  const rows = buildCsvTemplate()
    .map((row) => row.map((cell) => escapeCsv(cell)).join(","))
    .join("\n");
  const blob = new Blob([rows], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = "plantilla-padel.csv";
  link.click();
  URL.revokeObjectURL(url);
}

function formatMinutes(minutes: number) {
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;

  return `${String(hours).padStart(2, "0")}:${String(mins).padStart(2, "0")}`;
}

function minutesFromTime(time: string) {
  const [hours, mins] = time.split(":").map((part) => Number.parseInt(part, 10));

  return (Number.isFinite(hours) ? hours : 0) * 60 +
    (Number.isFinite(mins) ? mins : 0);
}

function dayLabel(day: DayKey) {
  if (day === "jueves") return "Jueves";
  if (day === "viernes") return "Viernes";

  return "Sabado";
}

function scheduleSlotSignature(schedule: Pick<ScheduleSlot, "court" | "day" | "time">) {
  return `${schedule.day}-${schedule.time}-${schedule.court}`;
}

function scheduleToManualOverride(
  schedule?: ScheduleAssignment,
): ManualScheduleOverride {
  return schedule
    ? {
        court: schedule.court,
        day: schedule.day,
        time: schedule.time,
      }
    : {
        court: 1,
        day: "sabado",
        time: "08:30",
      };
}

function manualOverrideToAssignment(
  override: ManualScheduleOverride,
): ScheduleAssignment {
  return {
    court: override.court,
    day: override.day,
    dayLabel: dayLabel(override.day),
    manual: true,
    minutes: minutesFromTime(override.time),
    slotIndex: -1,
    time: override.time,
  };
}

function makeSlotsForRange(
  day: DayKey,
  dayLabel: string,
  startMinutes: number,
  endMinutes: number,
) {
  const slots: ScheduleSlot[] = [];

  for (
    let minutes = startMinutes;
    minutes + 60 <= endMinutes;
    minutes += 60
  ) {
    for (let court = 1; court <= courtCount; court += 1) {
      slots.push({
        court,
        day,
        dayLabel,
        time: formatMinutes(minutes),
        minutes,
        slotIndex: -1,
      });
    }
  }

  return slots;
}

function tournamentSlots() {
  return [
    ...makeSlotsForRange("jueves", "Jueves", 10 * 60 + 30, 13 * 60 + 30),
    ...makeSlotsForRange("jueves", "Jueves", 17 * 60, 22 * 60 + 30),
    ...makeSlotsForRange("viernes", "Viernes", 10 * 60 + 30, 13 * 60 + 30),
    ...makeSlotsForRange("viernes", "Viernes", 17 * 60, 22 * 60 + 30),
    ...makeSlotsForRange("sabado", "Sabado", 8 * 60 + 30, 21 * 60 + 30),
  ].map((slot, slotIndex) => ({ ...slot, slotIndex }));
}

function parseRestrictionTime(hourText: string, minuteText?: string, suffix?: string) {
  let hours = Number.parseInt(hourText, 10);
  const minutes = Number.parseInt(minuteText || "0", 10);
  const normalizedSuffix = suffix?.toLowerCase();

  if (normalizedSuffix === "pm" && hours < 12) hours += 12;
  if (normalizedSuffix === "am" && hours === 12) hours = 0;

  return hours * 60 + minutes;
}

function mentionedRestrictionDays(text: string) {
  const days: DayKey[] = [];

  if (text.includes("jueves")) days.push("jueves");
  if (text.includes("viernes")) days.push("viernes");
  if (text.includes("sabado")) days.push("sabado");

  return days;
}

function dayIsDenied(text: string, day: DayKey) {
  const dayBeforeNo = new RegExp(
    `${day}\\s+(?:no|ni)(?!\\s+(?:entre|manana|tarde|puede))`,
  );
  const deniedBeforeDay = new RegExp(`(?:ni|sin)\\s+${day}`);

  return dayBeforeNo.test(text) || deniedBeforeDay.test(text);
}

function extractMinimumTimes(text: string) {
  const times: number[] = [];
  const plusPattern = /\+\s*(\d{1,2})(?::?(\d{2}))?\s*(am|pm)?/gi;
  const afterPattern =
    /(?:a partir de|despues de|desde)\s+(?:las\s+)?(\d{1,2})(?::?(\d{2}))?\s*(am|pm)?/gi;

  for (const match of text.matchAll(plusPattern)) {
    times.push(parseRestrictionTime(match[1], match[2], match[3]));
  }

  for (const match of text.matchAll(afterPattern)) {
    times.push(parseRestrictionTime(match[1], match[2], match[3]));
  }

  return times;
}

function extractMaximumTimes(text: string) {
  const times: number[] = [];
  const minusPattern = /-\s*(\d{1,2})(?::?(\d{2}))?\s*(am|pm)?/gi;
  const beforePattern =
    /(?:antes de|hasta)\s+(?:las\s+)?(\d{1,2})(?::?(\d{2}))?\s*(am|pm)?/gi;

  for (const match of text.matchAll(minusPattern)) {
    times.push(parseRestrictionTime(match[1], match[2], match[3]));
  }

  for (const match of text.matchAll(beforePattern)) {
    times.push(parseRestrictionTime(match[1], match[2], match[3]));
  }

  return times;
}

function deniedInterval(text: string) {
  if (!text.includes("no") && !text.includes("entre")) return null;

  const match = text.match(
    /entre\s+(?:las\s+)?(\d{1,2})(?::?(\d{2}))?\s*(am|pm)?\s+(?:y|a)\s+(?:las\s+)?(\d{1,2})(?::?(\d{2}))?\s*(am|pm)?/,
  );

  if (!match) return null;

  return {
    from: parseRestrictionTime(match[1], match[2], match[3]),
    to: parseRestrictionTime(match[4], match[5], match[6]),
  };
}

function slotMatchesRestriction(restriction: string | undefined, slot: ScheduleSlot) {
  const text = normalizeText(restriction || "");

  if (!text || text.includes("cualquier hora") || text.includes("a tope")) {
    return true;
  }

  const mentionedDays = mentionedRestrictionDays(text);
  const appliesToSlotDay =
    mentionedDays.length === 0 || mentionedDays.includes(slot.day);

  if (dayIsDenied(text, slot.day)) return false;

  if (appliesToSlotDay) {
    const morningDenied =
      text.includes("manana") &&
      (text.includes(" no") || text.includes("no ") || text.includes("ni "));
    const afternoonDenied =
      text.includes("tarde") &&
      (text.includes(" no") || text.includes("no ") || text.includes("ni "));

    if (morningDenied && slot.minutes < 14 * 60) return false;
    if (afternoonDenied && slot.minutes >= 14 * 60) return false;

    const interval = deniedInterval(text);
    if (
      interval &&
      slot.minutes >= interval.from &&
      slot.minutes < interval.to
    ) {
      return false;
    }

    const minimumTimes = extractMinimumTimes(text);
    const maximumTimes = extractMaximumTimes(text);
    const minimum = minimumTimes.length ? Math.min(...minimumTimes) : null;
    const maximum = maximumTimes.length ? Math.max(...maximumTimes) : null;

    if (minimum !== null && maximum !== null) {
      return slot.minutes >= minimum || slot.minutes <= maximum;
    }

    if (minimum !== null && slot.minutes < minimum) return false;
    if (maximum !== null && slot.minutes > maximum) return false;
  }

  return true;
}

function matchNeedsSchedule(match: Match) {
  return Boolean(match.sideA && match.sideB);
}

function matchFitsSlot(match: Match, slot: ScheduleSlot) {
  const restrictions = [match.sideA, match.sideB]
    .filter((team): team is Team => Boolean(team && !team.isPlaceholder))
    .map((team) => team.restriction)
    .filter(Boolean);

  return restrictions.every((restriction) =>
    slotMatchesRestriction(restriction, slot),
  );
}

function chronologicalMatchGroups(mainDraw: Draw, consolationDraw: Draw) {
  const groups: Match[][] = [];

  for (let index = 0; index < mainDraw.rounds.length; index += 1) {
    const stageMatches = [
      ...(mainDraw.rounds[index]?.matches ?? []),
      ...(index > 0 ? (consolationDraw.rounds[index - 1]?.matches ?? []) : []),
    ].filter(matchNeedsSchedule);

    if (stageMatches.length) groups.push(stageMatches);
  }

  for (
    let index = Math.max(0, mainDraw.rounds.length - 1);
    index < consolationDraw.rounds.length;
    index += 1
  ) {
    const consolationOnly = (consolationDraw.rounds[index]?.matches ?? []).filter(
      matchNeedsSchedule,
    );

    if (consolationOnly.length) {
      groups.push(consolationOnly);
    }
  }

  return groups;
}

function scheduleTimeGroups(slots: ScheduleSlot[]) {
  const groups: ScheduleTimeGroup[] = [];

  slots.forEach((slot) => {
    const current = groups.at(-1);

    if (
      current &&
      current.slots[0]?.day === slot.day &&
      current.slots[0]?.time === slot.time
    ) {
      current.slots.push(slot);
    } else {
      groups.push({
        groupIndex: groups.length,
        slots: [slot],
      });
    }
  });

  return groups;
}

type SchedulableMatch = {
  categoryId: string;
  match: Match;
};

function matchScheduleKey(categoryId: string, matchId: string) {
  return `${categoryId}::${matchId}`;
}

function globalChronologicalMatchGroups(drawSets: CategoryDrawSet[]) {
  const categoryGroups = drawSets.map((drawSet) => ({
    categoryId: drawSet.categoryId,
    groups: chronologicalMatchGroups(drawSet.mainDraw, drawSet.consolationDraw),
  }));
  const totalStages = Math.max(
    0,
    ...categoryGroups.map((drawSet) => drawSet.groups.length),
  );
  const groups: SchedulableMatch[][] = [];

  for (let index = 0; index < totalStages; index += 1) {
    const stageMatches = categoryGroups.flatMap((drawSet) =>
      (drawSet.groups[index] ?? []).map((match) => ({
        categoryId: drawSet.categoryId,
        match,
      })),
    );

    if (stageMatches.length) groups.push(stageMatches);
  }

  return groups;
}

function buildGlobalSchedule(
  drawSets: CategoryDrawSet[],
  manualScheduleOverrides: ManualScheduleMap = {},
): ScheduleResult {
  const groups = globalChronologicalMatchGroups(drawSets);
  const slots = tournamentSlots();
  const timeGroups = scheduleTimeGroups(slots);
  const assignments: Record<string, ScheduleAssignment> = {};
  const matchesByKey = new Map<string, Match>();
  const usedSlotIndexes = new Set<number>();
  let latestGroup = timeGroups.length - 1;
  let unscheduledConflicts = 0;

  groups.forEach((group) => {
    group.forEach((item) => {
      matchesByKey.set(matchScheduleKey(item.categoryId, item.match.id), item.match);
    });
  });

  groups
    .slice()
    .reverse()
    .forEach((group) => {
      let earliestUsedGroup = latestGroup;

      group
        .slice()
        .reverse()
        .forEach((item) => {
          const compatibleGroupIndex = timeGroups.findLastIndex(
            (timeGroup) =>
              timeGroup.groupIndex >= 0 &&
              timeGroup.groupIndex <= latestGroup &&
              timeGroup.slots.some(
                (slot) =>
                  !usedSlotIndexes.has(slot.slotIndex) &&
                  matchFitsSlot(item.match, slot),
              ),
          );
          const fallbackGroupIndex =
            compatibleGroupIndex >= 0
              ? compatibleGroupIndex
              : timeGroups.findLastIndex(
                  (timeGroup) =>
                    timeGroup.groupIndex >= 0 &&
                    timeGroup.groupIndex <= latestGroup &&
                    timeGroup.slots.some(
                      (slot) => !usedSlotIndexes.has(slot.slotIndex),
                    ),
                );
          const timeGroup =
            fallbackGroupIndex >= 0 ? timeGroups[fallbackGroupIndex] : null;
          const compatibleSlot = timeGroup?.slots.find(
            (slot) =>
              !usedSlotIndexes.has(slot.slotIndex) &&
              matchFitsSlot(item.match, slot),
          );
          const fallbackSlot =
            compatibleSlot ||
            timeGroup?.slots.find(
              (slot) => !usedSlotIndexes.has(slot.slotIndex),
            );

          if (!fallbackSlot || !timeGroup) {
            unscheduledConflicts += 1;
            return;
          }

          assignments[matchScheduleKey(item.categoryId, item.match.id)] = {
            ...fallbackSlot,
            conflict: !compatibleSlot,
          };

          usedSlotIndexes.add(fallbackSlot.slotIndex);
          earliestUsedGroup = Math.min(earliestUsedGroup, timeGroup.groupIndex);
        });

      latestGroup = earliestUsedGroup - 1;
    });

  Object.entries(manualScheduleOverrides).forEach(([key, override]) => {
    if (!matchesByKey.has(key)) return;

    assignments[key] = manualOverrideToAssignment(override);
  });

  const conflictKeys = new Set<string>();
  const occupiedSlots = new Map<string, string[]>();

  Object.entries(assignments).forEach(([key, assignment]) => {
    const match = matchesByKey.get(key);
    const slotKey = scheduleSlotSignature(assignment);
    const slotMatches = occupiedSlots.get(slotKey) ?? [];

    slotMatches.push(key);
    occupiedSlots.set(slotKey, slotMatches);

    if (assignment.conflict || (match && !matchFitsSlot(match, assignment))) {
      conflictKeys.add(key);
    }
  });

  occupiedSlots.forEach((keys) => {
    if (keys.length <= 1) return;

    keys.forEach((key) => conflictKeys.add(key));
  });

  conflictKeys.forEach((key) => {
    assignments[key] = {
      ...assignments[key],
      conflict: true,
    };
  });

  return {
    assignments,
    conflicts: unscheduledConflicts + conflictKeys.size,
    saturdayCount: Object.values(assignments).filter(
      (assignment) => assignment.day === "sabado",
    ).length,
    total: Object.keys(assignments).length,
  };
}

function scheduleLabel(schedule?: ScheduleAssignment) {
  if (!schedule) return "Sin horario";

  return `${schedule.dayLabel} ${schedule.time} · Pista ${schedule.court}`;
}

export default function TournamentAdmin() {
  const isClientReady = useClientReady();

  if (!isClientReady) {
    return (
      <main className="min-h-screen bg-[#f5f4ef] text-[#111816]">
        <div className="mx-auto flex w-full max-w-[1500px] flex-col gap-6 px-4 py-5 sm:px-6 lg:px-8">
          <AdminHeaderShell />
          <EmptyState />
        </div>
      </main>
    );
  }

  return (
    <AdminPasswordGate>
      <TournamentAdminClient initialState={readInitialAdminState()} />
    </AdminPasswordGate>
  );
}

function AdminPasswordGate({ children }: { children: ReactNode }) {
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [isAllowed, setIsAllowed] = useState(
    () => window.sessionStorage.getItem(adminAuthStorageKey) === "ok",
  );

  function handleSubmit(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();

    if (password === adminPassword) {
      window.sessionStorage.setItem(adminAuthStorageKey, "ok");
      setIsAllowed(true);
      setError("");
      return;
    }

    setError("Password incorrecta.");
  }

  if (isAllowed) return children;

  return (
    <main className="min-h-screen bg-[#f5f4ef] text-[#111816]">
      <div className="mx-auto flex min-h-screen w-full max-w-[980px] items-center px-4 py-8 sm:px-6 lg:px-8">
        <section className="w-full rounded-xl border border-black/10 bg-white p-6 shadow-xl shadow-black/10 sm:p-8">
          <div className="flex flex-col gap-6 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <div className="flex h-14 w-14 items-center justify-center rounded-xl bg-[#07110d] text-[#c59b45]">
                <LockKeyhole className="h-7 w-7" />
              </div>
              <p className="mt-5 text-xs font-bold uppercase tracking-[0.24em] text-[#9b7732]">
                Panel admin
              </p>
              <h1 className="mt-2 text-3xl font-black tracking-normal sm:text-5xl">
                Acceso privado
              </h1>
            </div>

            <form className="w-full max-w-sm" onSubmit={handleSubmit}>
              <label
                className="text-xs font-black uppercase tracking-[0.18em] text-black/45"
                htmlFor="admin-password"
              >
                Password
              </label>
              <input
                autoFocus
                autoComplete="current-password"
                className="mt-2 h-12 w-full rounded-lg border border-black/10 bg-[#f7f7f3] px-4 text-base font-bold outline-none transition focus:border-[#0f6b4b] focus:bg-white"
                id="admin-password"
                onChange={(event) => {
                  setPassword(event.target.value);
                  setError("");
                }}
                type="password"
                value={password}
              />
              {error ? (
                <p className="mt-2 text-sm font-bold text-red-700">{error}</p>
              ) : null}
              <button
                className="mt-4 inline-flex h-11 w-full items-center justify-center rounded-lg bg-[#0f6b4b] px-4 text-sm font-bold text-white transition hover:bg-[#11835b]"
                type="submit"
              >
                Entrar
              </button>
            </form>
          </div>
        </section>
      </div>
    </main>
  );
}

function TournamentAdminClient({
  initialState,
}: {
  initialState: InitialAdminState;
}) {
  const [previewCategories, setPreviewCategories] = useState<CategoryData[]>([]);
  const [previewCategoryId, setPreviewCategoryId] = useState("");
  const [categories, setCategories] = useState<CategoryData[]>(
    () => initialState.categories,
  );
  const [selectedCategoryId, setSelectedCategoryId] = useState(
    () => initialState.selectedCategoryId,
  );
  const [mainSelectionsByCategory, setMainSelectionsByCategory] =
    useState<CategorySelectionMap>(
      () => initialState.mainSelectionsByCategory,
    );
  const [consolationSelectionsByCategory, setConsolationSelectionsByCategory] =
    useState<CategorySelectionMap>(
      () => initialState.consolationSelectionsByCategory,
    );
  const [manualScheduleOverrides, setManualScheduleOverrides] =
    useState<ManualScheduleMap>(() => initialState.manualScheduleOverrides);
  const [importError, setImportError] = useState("");
  const [publishedSlug, setPublishedSlug] = useState(
    () => initialState.publishedSlug,
  );
  const [copiedPublicLink, setCopiedPublicLink] = useState(false);
  const [activeAdminTab, setActiveAdminTab] = useState<AdminTab>("draws");

  const activePreview = useMemo(
    () =>
      previewCategories.find((category) => category.id === previewCategoryId) ||
      previewCategories[0],
    [previewCategories, previewCategoryId],
  );

  const activeCategory = useMemo(
    () =>
      categories.find((category) => category.id === selectedCategoryId) ||
      categories[0],
    [categories, selectedCategoryId],
  );

  const drawSets = useMemo(
    () =>
      buildDrawSets(
        categories,
        mainSelectionsByCategory,
        consolationSelectionsByCategory,
      ),
    [categories, consolationSelectionsByCategory, mainSelectionsByCategory],
  );

  const activeDrawSet =
    drawSets.find((drawSet) => drawSet.categoryId === activeCategory?.id) ||
    drawSets[0];
  const mainDraw = activeDrawSet?.mainDraw ?? emptyMainDraw;
  const consolationDraw = activeDrawSet?.consolationDraw ?? emptyConsolationDraw;
  const schedule = useMemo(
    () => buildGlobalSchedule(drawSets, manualScheduleOverrides),
    [drawSets, manualScheduleOverrides],
  );
  const summaryRows = useMemo(
    () => buildScheduleSummaryRows(categories, drawSets, schedule.assignments),
    [categories, drawSets, schedule.assignments],
  );

  const firstRoundResolved = useMemo(() => {
    const matches = mainDraw.rounds[0]?.matches ?? [];

    return matches.filter((match) => match.loser).length;
  }, [mainDraw]);

  const publicPath = publishedSlug ? `/publico/${publishedSlug}` : "";
  const publicOrigin =
    typeof window === "undefined" ? "" : window.location.origin;
  const publicUrl = publicPath ? `${publicOrigin}${publicPath}` : "";

  useEffect(() => {
    if (!categories.length) {
      removeStorageItem(draftStorageKey);
      return;
    }

    const snapshot = buildTournamentSnapshot({
      categories,
      consolationSelectionsByCategory,
      mainSelectionsByCategory,
      manualScheduleOverrides,
      selectedCategoryId: activeCategory?.id || selectedCategoryId,
    });
    const draft: TournamentDraft = {
      ...snapshot,
      publishedSlug: publishedSlug || undefined,
      savedAt: new Date().toISOString(),
    };

    writeStorageJson(draftStorageKey, draft);

    if (publishedSlug) {
      savePublishedTournament(publishedSlug, snapshot);
    }
  }, [
    activeCategory?.id,
    categories,
    consolationSelectionsByCategory,
    mainSelectionsByCategory,
    manualScheduleOverrides,
    publishedSlug,
    selectedCategoryId,
  ]);

  async function handleUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";

    if (!file) return;

    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const firstSheet = workbook.SheetNames[0];
      const worksheet = firstSheet ? workbook.Sheets[firstSheet] : null;

      if (!worksheet) {
        setImportError("No se ha encontrado ninguna hoja en el archivo.");
        return;
      }

      const rows = XLSX.utils.sheet_to_json<SheetCell[]>(worksheet, {
        header: 1,
        defval: "",
      }) as SheetCell[][];
      const parsed = parseCategories(rows);

      if (!parsed.length) {
        setImportError("No se han encontrado parejas en columnas B y C.");
        return;
      }

      setPreviewCategories(parsed);
      setPreviewCategoryId(parsed[0].id);
      setImportError("");
    } catch {
      setImportError("No se ha podido leer el Excel.");
    }
  }

  function validatePreview() {
    if (!previewCategories.length) return;

    setCategories(previewCategories);
    setSelectedCategoryId(activePreview?.id || previewCategories[0].id);
    setPreviewCategories([]);
    setPreviewCategoryId("");
    setMainSelectionsByCategory({});
    setConsolationSelectionsByCategory({});
    setManualScheduleOverrides({});
  }

  function selectCategory(id: string) {
    setSelectedCategoryId(id);
  }

  function updatePair(
    categoryId: string,
    pairId: string,
    patch: Partial<Pick<Pair, "comment" | "playerOne" | "playerTwo" | "restriction" | "seed">>,
  ) {
    setCategories((current) =>
      current.map((category) =>
        category.id === categoryId
          ? {
              ...category,
              pairs: category.pairs.map((pair) =>
                pair.id === pairId
                  ? {
                      ...pair,
                      ...patch,
                    }
                  : pair,
              ),
            }
          : category,
      ),
    );
  }

  function clearCategoryResults(categoryId: string) {
    setMainSelectionsByCategory((current) => ({
      ...current,
      [categoryId]: {},
    }));
    setConsolationSelectionsByCategory((current) => ({
      ...current,
      [categoryId]: {},
    }));
  }

  function clearCategoryScheduleOverrides(categoryId: string) {
    setManualScheduleOverrides((current) =>
      withoutCategoryManualSchedules(current, categoryId),
    );
  }

  function randomizeCategoryBracket(categoryId: string) {
    const automaticOverrides = withoutCategoryManualSchedules(
      manualScheduleOverrides,
      categoryId,
    );
    const clearedMainSelections = {
      ...mainSelectionsByCategory,
      [categoryId]: {},
    };
    const clearedConsolationSelections = {
      ...consolationSelectionsByCategory,
      [categoryId]: {},
    };

    setCategories((current) => {
      const targetCategory = current.find((category) => category.id === categoryId);

      if (!targetCategory || targetCategory.pairs.length < 2) return current;

      let bestCategories = current;
      let bestScore = Number.POSITIVE_INFINITY;

      for (let attempt = 0; attempt < 120; attempt += 1) {
        const shuffledPairs = shufflePairs(targetCategory.pairs);
        const candidateCategories = current.map((category) =>
          category.id === categoryId
            ? {
                ...category,
                pairs: shuffledPairs,
              }
            : category,
        );
        const candidateDrawSets = buildDrawSets(
          candidateCategories,
          clearedMainSelections,
          clearedConsolationSelections,
        );
        const candidateSchedule = buildGlobalSchedule(
          candidateDrawSets,
          automaticOverrides,
        );
        const score =
          candidateSchedule.conflicts * 1000 - candidateSchedule.saturdayCount;

        if (score < bestScore) {
          bestScore = score;
          bestCategories = candidateCategories;
        }
      }

      return bestCategories;
    });
    clearCategoryResults(categoryId);
    setManualScheduleOverrides(automaticOverrides);
  }

  function addPair(categoryId: string) {
    setCategories((current) =>
      current.map((category) => {
        if (category.id !== categoryId) return category;

        const nextSeed = category.pairs.length + 1;

        return {
          ...category,
          pairs: [
            ...category.pairs,
            {
              category: category.name,
              comment: "",
              id: `pair-manual-${categoryId}-${Date.now().toString(36)}`,
              playerOne: `Jugador ${nextSeed}A`,
              playerTwo: `Jugador ${nextSeed}B`,
              restriction: "",
              seed: nextSeed,
            },
          ],
        };
      }),
    );
    clearCategoryResults(categoryId);
    clearCategoryScheduleOverrides(categoryId);
  }

  function removePair(categoryId: string, pairId: string) {
    setCategories((current) =>
      current.map((category) =>
        category.id === categoryId
          ? {
              ...category,
              pairs: category.pairs
                .filter((pair) => pair.id !== pairId)
                .map((pair, index) => ({
                  ...pair,
                  seed: index + 1,
                })),
            }
          : category,
      ),
    );
    clearCategoryResults(categoryId);
    clearCategoryScheduleOverrides(categoryId);
  }

  function movePairInBracket(categoryId: string, slotIndex: number, pairId: string) {
    setCategories((current) =>
      current.map((category) => {
        if (category.id !== categoryId) return category;

        const selectedIndex = category.pairs.findIndex((pair) => pair.id === pairId);

        if (selectedIndex < 0 || selectedIndex === slotIndex) return category;

        const nextPairs = [...category.pairs];
        const selectedPair = nextPairs[selectedIndex];
        const currentPair = nextPairs[slotIndex];

        nextPairs[slotIndex] = selectedPair;
        nextPairs[selectedIndex] = currentPair;

        return {
          ...category,
          pairs: nextPairs,
        };
      }),
    );
    clearCategoryResults(categoryId);
  }

  function updateManualSchedule(
    scheduleKey: string,
    override: ManualScheduleOverride,
  ) {
    setManualScheduleOverrides((current) => ({
      ...current,
      [scheduleKey]: override,
    }));
  }

  function resetManualSchedule(scheduleKey: string) {
    setManualScheduleOverrides((current) => {
      const next = { ...current };

      delete next[scheduleKey];

      return next;
    });
  }

  function selectMainWinner(match: Match, team: Team) {
    const categoryId = activeCategory?.id;

    if (!categoryId) return;
    if (!isPlayable(match.sideA) || !isPlayable(match.sideB)) return;

    setMainSelectionsByCategory((current) => ({
      ...current,
      [categoryId]: {
        ...(current[categoryId] ?? {}),
        [match.id]: team.id,
      },
    }));
    setConsolationSelectionsByCategory((current) => ({
      ...current,
      [categoryId]: {},
    }));
  }

  function selectConsolationWinner(match: Match, team: Team) {
    const categoryId = activeCategory?.id;

    if (!categoryId) return;
    if (!isPlayable(match.sideA) || !isPlayable(match.sideB)) return;

    setConsolationSelectionsByCategory((current) => ({
      ...current,
      [categoryId]: {
        ...(current[categoryId] ?? {}),
        [match.id]: team.id,
      },
    }));
  }

  function publishTournament() {
    if (!categories.length) return;

    const slug = publishedSlug || makeTournamentSlug();
    const snapshot = buildTournamentSnapshot({
      categories,
      consolationSelectionsByCategory,
      mainSelectionsByCategory,
      manualScheduleOverrides,
      selectedCategoryId: activeCategory?.id || selectedCategoryId,
    });

    savePublishedTournament(slug, snapshot);
    setPublishedSlug(slug);
    setCopiedPublicLink(false);
  }

  async function copyPublicLink() {
    if (!publicUrl) return;

    await navigator.clipboard.writeText(publicUrl);
    setCopiedPublicLink(true);
  }

  function resetAll() {
    setPreviewCategories([]);
    setPreviewCategoryId("");
    setCategories([]);
    setSelectedCategoryId("");
    setMainSelectionsByCategory({});
    setConsolationSelectionsByCategory({});
    setManualScheduleOverrides({});
    setImportError("");
    setPublishedSlug("");
    setCopiedPublicLink(false);
    removeStorageItem(draftStorageKey);
    removeStorageItem(activePublicationStorageKey);
  }

  return (
    <main className="min-h-screen bg-[#f5f4ef] text-[#111816]">
      <div className="mx-auto flex w-full max-w-[1500px] flex-col gap-6 px-4 py-5 sm:px-6 lg:px-8">
        <header className="rounded-xl bg-[#07110d] p-5 text-white shadow-xl shadow-black/10 sm:p-7">
          <div className="flex flex-col gap-5 lg:flex-row lg:items-end lg:justify-between">
            <div>
              <p className="text-xs font-bold uppercase tracking-[0.24em] text-[#c59b45]">
                Padel bracket
              </p>
              <h1 className="mt-2 text-3xl font-black tracking-normal sm:text-5xl">
                Cuadros del torneo
              </h1>
            </div>

            <div className="grid grid-cols-2 gap-3 sm:flex">
              <label className="inline-flex h-11 cursor-pointer items-center justify-center gap-2 rounded-lg bg-[#c59b45] px-4 text-sm font-bold text-black transition hover:bg-[#e0bd68]">
                <Upload className="h-4 w-4" />
                Subir Excel
                <input
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                  onChange={handleUpload}
                  type="file"
                />
              </label>
              <button
                className="inline-flex h-11 items-center justify-center gap-2 rounded-lg bg-[#0f6b4b] px-4 text-sm font-bold text-white transition hover:bg-[#11835b] disabled:cursor-not-allowed disabled:bg-white/10 disabled:text-white/40"
                disabled={!categories.length}
                onClick={publishTournament}
                type="button"
              >
                <Share2 className="h-4 w-4" />
                {publishedSlug ? "Actualizar publico" : "Publicar"}
              </button>
              {publishedSlug ? (
                <a
                  className="inline-flex h-11 items-center justify-center gap-2 rounded-lg border border-white/20 px-4 text-sm font-bold text-white transition hover:border-[#c59b45] hover:text-[#f2d081]"
                  href={publicPath}
                  rel="noreferrer"
                  target="_blank"
                >
                  <ExternalLink className="h-4 w-4" />
                  Ver publico
                </a>
              ) : null}
              <button
                className="inline-flex h-11 items-center justify-center gap-2 rounded-lg border border-white/20 px-4 text-sm font-bold text-white transition hover:border-[#c59b45] hover:text-[#f2d081]"
                onClick={downloadTemplate}
                type="button"
              >
                <FileSpreadsheet className="h-4 w-4" />
                Plantilla
              </button>
              <button
                className="inline-flex h-11 items-center justify-center gap-2 rounded-lg border border-white/20 px-4 text-sm font-bold text-white transition hover:border-[#c59b45] hover:text-[#f2d081]"
                onClick={resetAll}
                type="button"
              >
                <RotateCcw className="h-4 w-4" />
                Limpiar
              </button>
            </div>
          </div>
        </header>

        {importError ? (
          <Notice
            icon={AlertTriangle}
            tone="warning"
            title="Excel no valido"
            text={importError}
          />
        ) : null}

        {previewCategories.length ? (
          <PreviewPanel
            activeCategory={activePreview}
            categories={previewCategories}
            onCancel={() => setPreviewCategories([])}
            onSelect={setPreviewCategoryId}
            onValidate={validatePreview}
            selectedId={previewCategoryId}
          />
        ) : null}

        {categories.length ? (
          <section className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
            <StatCard
              icon={Users}
              label="Categorias"
              value={categories.length}
            />
            <StatCard
              icon={Trophy}
              label="Parejas"
              value={activeCategory?.pairs.length ?? 0}
            />
            <StatCard
              icon={Medal}
              label="En consolacion"
              value={firstRoundResolved}
            />
            <StatCard
              icon={CalendarClock}
              label="Sabado total"
              value={`${schedule.saturdayCount}/${schedule.total}`}
            />
          </section>
        ) : null}

        {categories.length ? (
          <PublishPanel
            copied={copiedPublicLink}
            onCopy={copyPublicLink}
            onPublish={publishTournament}
            publicPath={publicPath}
            publicUrl={publicUrl}
            published={Boolean(publishedSlug)}
          />
        ) : null}

        {categories.length ? (
          <AdminViewTabs activeTab={activeAdminTab} onChange={setActiveAdminTab} />
        ) : null}

        {categories.length && activeAdminTab === "summary" ? (
          <ScheduleSummaryPanel rows={summaryRows} />
        ) : null}

        {categories.length && activeAdminTab === "draws" ? (
          <>
            <section className="rounded-xl border border-black/10 bg-white p-4 shadow-sm">
              <div className="flex flex-wrap gap-2">
                {categories.map((category) => (
                  <button
                    className={`rounded-lg border px-3 py-2 text-sm font-bold transition ${
                      category.id === activeCategory?.id
                        ? "border-[#0f6b4b] bg-[#0f6b4b] text-white"
                        : "border-black/10 bg-[#f7f7f3] text-[#111816] hover:border-[#0f6b4b]"
                    }`}
                    key={category.id}
                    onClick={() => selectCategory(category.id)}
                    type="button"
                  >
                    {category.name}
                    <span className="ml-2 text-xs opacity-70">
                      {category.pairs.length}
                    </span>
                  </button>
                ))}
              </div>
            </section>

            {activeCategory ? (
              <TournamentEditPanel
                category={activeCategory}
                drawSet={activeDrawSet}
                manualScheduleOverrides={manualScheduleOverrides}
                onAddPair={addPair}
            onClearCategorySchedule={clearCategoryScheduleOverrides}
            onRemovePair={removePair}
            onResetManualSchedule={resetManualSchedule}
            onMovePairInBracket={movePairInBracket}
            onRandomizeBracket={randomizeCategoryBracket}
            onUpdateManualSchedule={updateManualSchedule}
            onUpdatePair={updatePair}
            schedule={schedule.assignments}
              />
            ) : null}
          </>
        ) : null}

        {categories.length && schedule.conflicts ? (
          <Notice
            icon={AlertTriangle}
            tone="warning"
            title="Revisar horarios"
            text={`${schedule.conflicts} partido(s) no encajan perfectamente con las restricciones y se han marcado en amarillo.`}
          />
        ) : null}

        {!categories.length ? (
          <EmptyState />
        ) : activeAdminTab === "draws" ? (
          <div className="grid gap-6">
            <BracketBoard
              categoryId={activeCategory?.id || ""}
              draw={mainDraw}
              onPick={selectMainWinner}
              schedule={schedule.assignments}
              subtitle="Selecciona el ganador de cada partido. El perdedor de primera ronda aparece abajo."
            />
            <BracketBoard
              categoryId={activeCategory?.id || ""}
              draw={consolationDraw}
              emptyText="La consolacion se llenara cuando marques perdedores en la primera ronda."
              onPick={selectConsolationWinner}
              schedule={schedule.assignments}
              subtitle="Cuadro secundario con los perdedores del primer partido."
            />
          </div>
        ) : null}
      </div>
    </main>
  );
}

function AdminHeaderShell() {
  return (
    <header className="rounded-xl bg-[#07110d] p-5 text-white shadow-xl shadow-black/10 sm:p-7">
      <div className="flex flex-col gap-5 lg:flex-row lg:items-end lg:justify-between">
        <div>
          <p className="text-xs font-bold uppercase tracking-[0.24em] text-[#c59b45]">
            Padel bracket
          </p>
          <h1 className="mt-2 text-3xl font-black tracking-normal sm:text-5xl">
            Cuadros del torneo
          </h1>
        </div>

        <div className="grid grid-cols-2 gap-3 sm:flex">
          <button
            className="inline-flex h-11 items-center justify-center gap-2 rounded-lg bg-white/10 px-4 text-sm font-bold text-white/40"
            disabled
            type="button"
          >
            <Upload className="h-4 w-4" />
            Subir Excel
          </button>
          <button
            className="inline-flex h-11 items-center justify-center gap-2 rounded-lg bg-white/10 px-4 text-sm font-bold text-white/40"
            disabled
            type="button"
          >
            <Share2 className="h-4 w-4" />
            Publicar
          </button>
        </div>
      </div>
    </header>
  );
}

function AdminViewTabs({
  activeTab,
  onChange,
}: {
  activeTab: AdminTab;
  onChange: (tab: AdminTab) => void;
}) {
  const tabs: { label: string; value: AdminTab }[] = [
    { label: "Cuadros", value: "draws" },
    { label: "Resumen", value: "summary" },
  ];

  return (
    <section className="rounded-xl border border-black/10 bg-white p-2 shadow-sm">
      <div className="grid grid-cols-2 gap-2 sm:flex">
        {tabs.map((tab) => (
          <button
            className={`h-10 rounded-lg px-4 text-sm font-black transition ${
              activeTab === tab.value
                ? "bg-[#0f6b4b] text-white"
                : "bg-[#f7f7f3] text-[#111816] hover:bg-[#e8f3ee]"
            }`}
            key={tab.value}
            onClick={() => onChange(tab.value)}
            type="button"
          >
            {tab.label}
          </button>
        ))}
      </div>
    </section>
  );
}

const scheduleDayOptions: { label: string; value: DayKey }[] = [
  { label: "Jueves", value: "jueves" },
  { label: "Viernes", value: "viernes" },
  { label: "Sabado", value: "sabado" },
];

const courtOptions = Array.from({ length: courtCount }, (_, index) => index + 1);

function timeOptionsForDay(day: DayKey) {
  return Array.from(
    new Set(
      tournamentSlots()
        .filter((slot) => slot.day === day)
        .map((slot) => slot.time),
    ),
  );
}

function editableMatchesFromDrawSet(categoryId: string, drawSet?: CategoryDrawSet) {
  if (!drawSet) return [];

  return [
    ...drawSet.mainDraw.rounds.flatMap((round) =>
      round.matches.map((match) => ({
        drawName: "Principal",
        match,
        roundName: round.name,
        scheduleKey: matchScheduleKey(categoryId, match.id),
      })),
    ),
    ...drawSet.consolationDraw.rounds.flatMap((round) =>
      round.matches.map((match) => ({
        drawName: "Consolacion",
        match,
        roundName: round.name,
        scheduleKey: matchScheduleKey(categoryId, match.id),
      })),
    ),
  ].filter(({ match }) => matchNeedsSchedule(match));
}

function matchTeamsLabel(match: Match) {
  return [match.sideA?.name || "BYE", match.sideB?.name || "BYE"].join(" vs ");
}

function firstRoundPairSlots(category: CategoryData) {
  const size = nextPowerOfTwo(category.pairs.length);
  const slots = Array<Pair | null>(size).fill(null);

  category.pairs.forEach((pair, index) => {
    slots[index] = pair;
  });

  return Array.from({ length: size / 2 }, (_, matchIndex) => ({
    matchLabel: `P${matchIndex + 1}`,
    sideA: slots[matchIndex * 2],
    sideAIndex: matchIndex * 2,
    sideB: slots[matchIndex * 2 + 1],
    sideBIndex: matchIndex * 2 + 1,
  }));
}

function shufflePairs(pairs: Pair[]) {
  const shuffled = [...pairs];

  for (let index = shuffled.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(Math.random() * (index + 1));
    const current = shuffled[index];

    shuffled[index] = shuffled[swapIndex];
    shuffled[swapIndex] = current;
  }

  return shuffled;
}

type ScheduleSummaryRow = {
  categoryName: string;
  conflict: boolean;
  court: number | null;
  day: DayKey | null;
  dayLabel: string;
  drawName: string;
  manual: boolean;
  matchId: string;
  minutes: number;
  roundName: string;
  scheduleKey: string;
  teams: string;
  time: string;
};

function buildScheduleSummaryRows(
  categories: CategoryData[],
  drawSets: CategoryDrawSet[],
  schedule: Record<string, ScheduleAssignment>,
) {
  const categoryById = new Map(categories.map((category) => [category.id, category]));
  const categoryOrder = new Map(
    categories.map((category, index) => [category.id, index]),
  );

  return drawSets
    .flatMap((drawSet) => {
      const category = categoryById.get(drawSet.categoryId);
      const categoryName = category?.name || "Categoria";
      const drawRows = [
        ...drawSet.mainDraw.rounds.flatMap((round) =>
          round.matches.map((match) => ({
            drawName: "Principal",
            match,
            roundName: round.name,
          })),
        ),
        ...drawSet.consolationDraw.rounds.flatMap((round) =>
          round.matches.map((match) => ({
            drawName: "Consolacion",
            match,
            roundName: round.name,
          })),
        ),
      ];

      return drawRows
        .filter(({ match }) => matchNeedsSchedule(match))
        .map(({ drawName, match, roundName }) => {
          const scheduleKey = matchScheduleKey(drawSet.categoryId, match.id);
          const assignment = schedule[scheduleKey];

          return {
            categoryName,
            conflict: Boolean(assignment?.conflict),
            court: assignment?.court ?? null,
            day: assignment?.day ?? null,
            dayLabel: assignment?.dayLabel ?? "Sin dia",
            drawName,
            manual: Boolean(assignment?.manual),
            matchId: match.id,
            minutes: assignment?.minutes ?? Number.MAX_SAFE_INTEGER,
            roundName,
            scheduleKey,
            teams: matchTeamsLabel(match),
            time: assignment?.time ?? "Sin hora",
            categorySort: categoryOrder.get(drawSet.categoryId) ?? 999,
          };
        });
    })
    .sort((left, right) => {
      const dayDiff =
        (left.day ? dayOrder[left.day] : 99) -
        (right.day ? dayOrder[right.day] : 99);

      if (dayDiff) return dayDiff;
      if (left.minutes !== right.minutes) return left.minutes - right.minutes;
      if ((left.court ?? 99) !== (right.court ?? 99)) {
        return (left.court ?? 99) - (right.court ?? 99);
      }
      if (left.categorySort !== right.categorySort) {
        return left.categorySort - right.categorySort;
      }

      return left.matchId.localeCompare(right.matchId);
    });
}

function summaryDayClass(day: DayKey | null) {
  if (day === "jueves") return "bg-sky-50 text-sky-900";
  if (day === "viernes") return "bg-violet-50 text-violet-900";
  if (day === "sabado") return "bg-emerald-50 text-emerald-900";

  return "bg-stone-100 text-stone-700";
}

function summaryDayPillClass(day: DayKey | null) {
  if (day === "jueves") return "bg-sky-200 text-sky-950";
  if (day === "viernes") return "bg-violet-200 text-violet-950";
  if (day === "sabado") return "bg-emerald-200 text-emerald-950";

  return "bg-stone-200 text-stone-800";
}

function summaryCategoryClass(categoryName: string) {
  const palette = [
    "bg-[#e8f3ee] text-[#0b553b]",
    "bg-[#fff3d6] text-[#7a5818]",
    "bg-sky-100 text-sky-900",
    "bg-rose-100 text-rose-900",
    "bg-violet-100 text-violet-900",
  ];
  const index = normalizeText(categoryName)
    .split("")
    .reduce((sum, char) => sum + char.charCodeAt(0), 0);

  return palette[index % palette.length];
}

function summaryDrawClass(drawName: string) {
  return drawName === "Principal"
    ? "bg-[#0f6b4b] text-white"
    : "bg-[#c59b45] text-black";
}

function ScheduleSummaryPanel({ rows }: { rows: ScheduleSummaryRow[] }) {
  return (
    <section className="rounded-xl border border-black/10 bg-white p-4 shadow-sm">
      <div className="mb-4 flex flex-col gap-2 sm:flex-row sm:items-end sm:justify-between">
        <div>
          <p className="text-xs font-bold uppercase tracking-[0.22em] text-[#0f6b4b]">
            Resumen
          </p>
          <h2 className="mt-1 text-2xl font-black">Listado de partidos</h2>
        </div>
        <div className="rounded-lg bg-[#f7f7f3] px-3 py-2 text-sm font-black text-black/55">
          {rows.length} partidos
        </div>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full min-w-[1060px] border-collapse text-sm">
          <thead>
            <tr className="bg-[#07110d] text-left text-white">
              <th className="px-3 py-3 font-bold">Dia</th>
              <th className="px-3 py-3 font-bold">Hora</th>
              <th className="px-3 py-3 font-bold">Pista</th>
              <th className="px-3 py-3 font-bold">Categoria</th>
              <th className="px-3 py-3 font-bold">Cuadro</th>
              <th className="px-3 py-3 font-bold">Ronda</th>
              <th className="px-3 py-3 font-bold">Partido</th>
              <th className="px-3 py-3 font-bold">Parejas</th>
              <th className="px-3 py-3 font-bold">Estado</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr
                className={`border-b border-black/10 ${
                  row.conflict ? "bg-amber-50" : summaryDayClass(row.day)
                }`}
                key={row.scheduleKey}
              >
                <td className="px-3 py-3">
                  <span
                    className={`rounded px-2 py-1 text-xs font-black ${summaryDayPillClass(row.day)}`}
                  >
                    {row.dayLabel}
                  </span>
                </td>
                <td className="px-3 py-3">
                  <span className="rounded bg-white/75 px-2 py-1 font-black">
                    {row.time}
                  </span>
                </td>
                <td className="px-3 py-3">
                  <span className="rounded bg-black/10 px-2 py-1 font-black">
                    {row.court ? `Pista ${row.court}` : "-"}
                  </span>
                </td>
                <td className="px-3 py-3">
                  <span
                    className={`rounded px-2 py-1 text-xs font-black ${summaryCategoryClass(row.categoryName)}`}
                  >
                    {row.categoryName}
                  </span>
                </td>
                <td className="px-3 py-3">
                  <span
                    className={`rounded px-2 py-1 text-xs font-black ${summaryDrawClass(row.drawName)}`}
                  >
                    {row.drawName}
                  </span>
                </td>
                <td className="px-3 py-3">{row.roundName}</td>
                <td className="px-3 py-3 font-black">{row.matchId}</td>
                <td className="px-3 py-3 font-bold">{row.teams}</td>
                <td className="px-3 py-3">
                  <span
                    className={`rounded px-2 py-1 text-xs font-black ${
                      row.conflict
                        ? "bg-amber-200 text-amber-950"
                        : row.manual
                          ? "bg-[#fff3d6] text-[#7a5818]"
                          : "bg-[#e8f3ee] text-[#0b553b]"
                    }`}
                  >
                    {row.conflict ? "Revisar" : row.manual ? "Manual" : "Auto"}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function TournamentEditPanel({
  category,
  drawSet,
  manualScheduleOverrides,
  onAddPair,
  onClearCategorySchedule,
  onMovePairInBracket,
  onRandomizeBracket,
  onRemovePair,
  onResetManualSchedule,
  onUpdateManualSchedule,
  onUpdatePair,
  schedule,
}: {
  category: CategoryData;
  drawSet?: CategoryDrawSet;
  manualScheduleOverrides: ManualScheduleMap;
  onAddPair: (categoryId: string) => void;
  onClearCategorySchedule: (categoryId: string) => void;
  onMovePairInBracket: (
    categoryId: string,
    slotIndex: number,
    pairId: string,
  ) => void;
  onRandomizeBracket: (categoryId: string) => void;
  onRemovePair: (categoryId: string, pairId: string) => void;
  onResetManualSchedule: (scheduleKey: string) => void;
  onUpdateManualSchedule: (
    scheduleKey: string,
    override: ManualScheduleOverride,
  ) => void;
  onUpdatePair: (
    categoryId: string,
    pairId: string,
    patch: Partial<
      Pick<Pair, "comment" | "playerOne" | "playerTwo" | "restriction" | "seed">
    >,
  ) => void;
  schedule: Record<string, ScheduleAssignment>;
}) {
  const editableMatches = editableMatchesFromDrawSet(category.id, drawSet);
  const bracketPairSlots = firstRoundPairSlots(category);
  const [showBracketEditor, setShowBracketEditor] = useState(false);

  function updateScheduleValue(
    scheduleKey: string,
    next: Partial<ManualScheduleOverride>,
  ) {
    const current = scheduleToManualOverride(
      manualScheduleOverrides[scheduleKey]
        ? manualOverrideToAssignment(manualScheduleOverrides[scheduleKey])
        : schedule[scheduleKey],
    );
    const nextDay = next.day ?? current.day;
    const dayTimes = timeOptionsForDay(nextDay);
    const nextTime =
      next.time && dayTimes.includes(next.time)
        ? next.time
        : dayTimes.includes(current.time)
          ? current.time
          : dayTimes[0];

    onUpdateManualSchedule(scheduleKey, {
      court: next.court ?? current.court,
      day: nextDay,
      time: nextTime,
    });
  }

  return (
    <section className="rounded-xl border border-black/10 bg-white p-4 shadow-sm">
      <div className="flex flex-col gap-2 sm:flex-row sm:items-end sm:justify-between">
        <div>
          <p className="text-xs font-bold uppercase tracking-[0.22em] text-[#0f6b4b]">
            Edicion manual
          </p>
          <h2 className="mt-1 text-2xl font-black">{category.name}</h2>
          <p className="mt-1 text-sm font-semibold text-black/55">
            Ajusta parejas, restricciones y horarios antes de publicar.
          </p>
        </div>
        <button
          className="inline-flex h-10 items-center justify-center gap-2 rounded-lg bg-[#111816] px-4 text-sm font-bold text-white transition hover:bg-black"
          onClick={() => onAddPair(category.id)}
          type="button"
        >
          <Plus className="h-4 w-4" />
          Anadir pareja
        </button>
      </div>

      <div className="mt-5 rounded-xl border border-[#0f6b4b]/20 bg-[#f7f7f3] p-3">
        <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <h3 className="text-sm font-black uppercase tracking-[0.18em] text-[#0b553b]">
              Emparejamientos
            </h3>
            <p className="mt-1 text-sm font-semibold text-black/55">
              Cambia las parejas de sitio en la primera ronda del cuadro.
            </p>
          </div>
          <div className="flex flex-col gap-2 sm:flex-row">
            <button
              className="inline-flex h-10 items-center justify-center gap-2 rounded-lg bg-[#0f6b4b] px-4 text-sm font-bold text-white transition hover:bg-[#11835b]"
              onClick={() => onRandomizeBracket(category.id)}
              type="button"
            >
              <RotateCcw className="h-4 w-4" />
              Random con horarios
            </button>
            <button
              className="inline-flex h-10 items-center justify-center rounded-lg border border-black/10 bg-white px-4 text-sm font-bold transition hover:border-[#0f6b4b]"
              onClick={() => setShowBracketEditor((current) => !current)}
              type="button"
            >
              {showBracketEditor ? "Ocultar" : "Cambiar brackets a mano"}
            </button>
          </div>
        </div>

        {showBracketEditor ? (
          <div className="mt-4 grid gap-3 lg:grid-cols-2">
            {bracketPairSlots.map((slot) => (
              <div
                className="rounded-lg border border-black/10 bg-white p-3"
                key={slot.matchLabel}
              >
                <p className="mb-3 text-xs font-black uppercase tracking-[0.16em] text-black/45">
                  {slot.matchLabel}
                </p>
                <div className="grid gap-2">
                  <BracketSlotSelect
                    category={category}
                    label="Pareja A"
                    onMovePair={onMovePairInBracket}
                    pair={slot.sideA}
                    slotIndex={slot.sideAIndex}
                  />
                  <BracketSlotSelect
                    category={category}
                    label="Pareja B"
                    onMovePair={onMovePairInBracket}
                    pair={slot.sideB}
                    slotIndex={slot.sideBIndex}
                  />
                </div>
              </div>
            ))}
          </div>
        ) : null}
      </div>

      <div className="mt-5 grid gap-6 xl:grid-cols-[minmax(0,1.15fr)_minmax(420px,0.85fr)]">
        <div className="min-w-0">
          <div className="mb-3 flex items-center justify-between gap-3">
            <h3 className="text-sm font-black uppercase tracking-[0.18em] text-black/45">
              Parejas
            </h3>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full min-w-[900px] border-collapse text-sm">
              <thead>
                <tr className="bg-[#f7f7f3] text-left text-xs uppercase tracking-[0.14em] text-black/45">
                  <th className="px-3 py-3 font-black">#</th>
                  <th className="px-3 py-3 font-black">Jugador 1</th>
                  <th className="px-3 py-3 font-black">Jugador 2</th>
                  <th className="px-3 py-3 font-black">Restriccion</th>
                  <th className="px-3 py-3 font-black">Comentario</th>
                  <th className="px-3 py-3 font-black">Quitar</th>
                </tr>
              </thead>
              <tbody>
                {category.pairs.map((pair) => (
                  <tr className="border-b border-black/10" key={pair.id}>
                    <td className="px-3 py-2">
                      <input
                        className="h-9 w-14 rounded-lg border border-black/10 bg-white px-2 text-sm font-bold outline-none focus:border-[#0f6b4b]"
                        min={1}
                        onChange={(event) =>
                          onUpdatePair(category.id, pair.id, {
                            seed:
                              Number.parseInt(event.target.value, 10) ||
                              pair.seed,
                          })
                        }
                        type="number"
                        value={pair.seed}
                      />
                    </td>
                    <td className="px-3 py-2">
                      <input
                        className="h-9 w-full rounded-lg border border-black/10 bg-white px-3 text-sm font-bold outline-none focus:border-[#0f6b4b]"
                        onBlur={(event) =>
                          onUpdatePair(category.id, pair.id, {
                            playerOne: properCaseName(event.target.value),
                          })
                        }
                        onChange={(event) =>
                          onUpdatePair(category.id, pair.id, {
                            playerOne: event.target.value,
                          })
                        }
                        value={pair.playerOne}
                      />
                    </td>
                    <td className="px-3 py-2">
                      <input
                        className="h-9 w-full rounded-lg border border-black/10 bg-white px-3 text-sm font-bold outline-none focus:border-[#0f6b4b]"
                        onBlur={(event) =>
                          onUpdatePair(category.id, pair.id, {
                            playerTwo: properCaseName(event.target.value),
                          })
                        }
                        onChange={(event) =>
                          onUpdatePair(category.id, pair.id, {
                            playerTwo: event.target.value,
                          })
                        }
                        value={pair.playerTwo}
                      />
                    </td>
                    <td className="px-3 py-2">
                      <input
                        className="h-9 w-full rounded-lg border border-black/10 bg-white px-3 text-sm outline-none focus:border-[#0f6b4b]"
                        onChange={(event) =>
                          onUpdatePair(category.id, pair.id, {
                            restriction: event.target.value,
                          })
                        }
                        value={pair.restriction}
                      />
                    </td>
                    <td className="px-3 py-2">
                      <input
                        className="h-9 w-full rounded-lg border border-black/10 bg-white px-3 text-sm outline-none focus:border-[#0f6b4b]"
                        onChange={(event) =>
                          onUpdatePair(category.id, pair.id, {
                            comment: event.target.value,
                          })
                        }
                        value={pair.comment}
                      />
                    </td>
                    <td className="px-3 py-2">
                      <button
                        className="inline-flex h-9 w-9 items-center justify-center rounded-lg border border-black/10 text-black/55 transition hover:border-red-300 hover:bg-red-50 hover:text-red-700"
                        onClick={() => onRemovePair(category.id, pair.id)}
                        type="button"
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div className="min-w-0">
          <div className="mb-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
            <h3 className="text-sm font-black uppercase tracking-[0.18em] text-black/45">
              Horarios
            </h3>
            <button
              className="inline-flex h-9 items-center justify-center gap-2 rounded-lg border border-black/10 px-3 text-sm font-bold transition hover:border-[#0f6b4b]"
              onClick={() => onClearCategorySchedule(category.id)}
              type="button"
            >
              <RotateCcw className="h-4 w-4" />
              Recalcular
            </button>
          </div>

          <div className="grid max-h-[560px] gap-3 overflow-auto pr-1">
            {editableMatches.map(({ drawName, match, roundName, scheduleKey }) => {
              const currentSchedule = schedule[scheduleKey];
              const manualOverride = manualScheduleOverrides[scheduleKey];
              const selectedSchedule = manualOverride
                ? manualOverrideToAssignment(manualOverride)
                : currentSchedule;
              const currentValue = scheduleToManualOverride(selectedSchedule);
              const timeOptions = timeOptionsForDay(currentValue.day);

              return (
                <div
                  className={`rounded-lg border p-3 ${
                    currentSchedule?.conflict
                      ? "border-amber-300 bg-amber-50"
                      : "border-black/10 bg-[#fbfbf8]"
                  }`}
                  key={scheduleKey}
                >
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0">
                      <p className="text-xs font-black uppercase tracking-[0.14em] text-black/45">
                        {drawName} · {roundName} · {match.id}
                      </p>
                      <p className="mt-1 line-clamp-2 text-sm font-black">
                        {matchTeamsLabel(match)}
                      </p>
                    </div>
                    <span
                      className={`shrink-0 rounded px-2 py-1 text-xs font-black ${
                        manualOverride
                          ? "bg-[#fff3d6] text-[#7a5818]"
                          : "bg-[#e8f3ee] text-[#0b553b]"
                      }`}
                    >
                      {manualOverride ? "Manual" : "Auto"}
                    </span>
                  </div>

                  <div className="mt-3 grid grid-cols-[1fr_1fr_84px_auto] gap-2">
                    <select
                      className="h-9 rounded-lg border border-black/10 bg-white px-2 text-sm font-bold outline-none focus:border-[#0f6b4b]"
                      onChange={(event) =>
                        updateScheduleValue(scheduleKey, {
                          day: event.target.value as DayKey,
                        })
                      }
                      value={currentValue.day}
                    >
                      {scheduleDayOptions.map((option) => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>

                    <select
                      className="h-9 rounded-lg border border-black/10 bg-white px-2 text-sm font-bold outline-none focus:border-[#0f6b4b]"
                      onChange={(event) =>
                        updateScheduleValue(scheduleKey, {
                          time: event.target.value,
                        })
                      }
                      value={currentValue.time}
                    >
                      {timeOptions.map((time) => (
                        <option key={time} value={time}>
                          {time}
                        </option>
                      ))}
                    </select>

                    <select
                      className="h-9 rounded-lg border border-black/10 bg-white px-2 text-sm font-bold outline-none focus:border-[#0f6b4b]"
                      onChange={(event) =>
                        updateScheduleValue(scheduleKey, {
                          court: Number.parseInt(event.target.value, 10),
                        })
                      }
                      value={currentValue.court}
                    >
                      {courtOptions.map((court) => (
                        <option key={court} value={court}>
                          P{court}
                        </option>
                      ))}
                    </select>

                    <button
                      className="inline-flex h-9 w-9 items-center justify-center rounded-lg border border-black/10 text-black/55 transition hover:border-[#0f6b4b] hover:text-[#0f6b4b] disabled:cursor-not-allowed disabled:opacity-35"
                      disabled={!manualOverride}
                      onClick={() => onResetManualSchedule(scheduleKey)}
                      type="button"
                    >
                      <RotateCcw className="h-4 w-4" />
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </section>
  );
}

function PublishPanel({
  copied,
  onCopy,
  onPublish,
  publicPath,
  publicUrl,
  published,
}: {
  copied: boolean;
  onCopy: () => void;
  onPublish: () => void;
  publicPath: string;
  publicUrl: string;
  published: boolean;
}) {
  return (
    <section className="rounded-xl border border-[#0f6b4b]/20 bg-white p-4 shadow-sm">
      <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <p className="text-xs font-bold uppercase tracking-[0.22em] text-[#0f6b4b]">
            Panel admin
          </p>
          <h2 className="mt-1 text-xl font-black">
            {published ? "Cuadro publicado" : "Publica cuando este listo"}
          </h2>
          <p className="mt-1 text-sm font-semibold text-black/55">
            Desde este panel puedes seguir marcando ganadores; el enlace publico
            se actualiza con tus cambios.
          </p>
        </div>

        <div className="flex flex-col gap-2 sm:flex-row">
          <button
            className="inline-flex h-10 items-center justify-center gap-2 rounded-lg bg-[#0f6b4b] px-4 text-sm font-bold text-white transition hover:bg-[#11835b]"
            onClick={onPublish}
            type="button"
          >
            <Share2 className="h-4 w-4" />
            {published ? "Actualizar publico" : "Publicar cuadro"}
          </button>
          {published ? (
            <a
              className="inline-flex h-10 items-center justify-center gap-2 rounded-lg border border-black/10 px-4 text-sm font-bold text-[#111816] transition hover:border-[#0f6b4b]"
              href={publicPath}
              rel="noreferrer"
              target="_blank"
            >
              <ExternalLink className="h-4 w-4" />
              Abrir
            </a>
          ) : null}
        </div>
      </div>

      {published ? (
        <div className="mt-4 flex flex-col gap-2 rounded-lg bg-[#f7f7f3] p-3 sm:flex-row sm:items-center">
          <code className="min-w-0 flex-1 truncate text-sm font-bold text-[#0b553b]">
            {publicUrl}
          </code>
          <button
            className="inline-flex h-9 items-center justify-center gap-2 rounded-lg bg-[#111816] px-3 text-sm font-bold text-white transition hover:bg-black"
            onClick={onCopy}
            type="button"
          >
            <Copy className="h-4 w-4" />
            {copied ? "Copiado" : "Copiar"}
          </button>
        </div>
      ) : null}
    </section>
  );
}

function BracketSlotSelect({
  category,
  label,
  onMovePair,
  pair,
  slotIndex,
}: {
  category: CategoryData;
  label: string;
  onMovePair: (categoryId: string, slotIndex: number, pairId: string) => void;
  pair: Pair | null;
  slotIndex: number;
}) {
  return (
    <label className="grid gap-1">
      <span className="text-xs font-black uppercase tracking-[0.14em] text-black/45">
        {label}
      </span>
      {pair ? (
        <select
          className="h-10 rounded-lg border border-black/10 bg-[#fbfbf8] px-3 text-sm font-bold outline-none transition focus:border-[#0f6b4b]"
          onChange={(event) =>
            onMovePair(category.id, slotIndex, event.target.value)
          }
          value={pair.id}
        >
          {category.pairs.map((option) => (
            <option key={option.id} value={option.id}>
              {option.seed}. {pairName(option)}
            </option>
          ))}
        </select>
      ) : (
        <div className="flex h-10 items-center rounded-lg border border-dashed border-black/15 bg-black/[0.03] px-3 text-sm font-black text-black/35">
          BYE
        </div>
      )}
    </label>
  );
}

function PreviewPanel({
  activeCategory,
  categories,
  onCancel,
  onSelect,
  onValidate,
  selectedId,
}: {
  activeCategory?: CategoryData;
  categories: CategoryData[];
  onCancel: () => void;
  onSelect: (id: string) => void;
  onValidate: () => void;
  selectedId: string;
}) {
  const rows = activeCategory?.pairs.slice(0, previewLimit) ?? [];

  return (
    <section className="rounded-xl border-2 border-[#c59b45] bg-white p-4 shadow-lg shadow-[#c59b45]/10">
      <div className="flex flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
        <div>
          <p className="text-xs font-bold uppercase tracking-[0.22em] text-[#9b7732]">
            Previsualizacion
          </p>
          <h2 className="mt-1 text-2xl font-black text-[#111816]">
            Asi se va a subir
          </h2>
        </div>

        <div className="flex flex-wrap gap-2">
          <button
            className="inline-flex h-10 items-center justify-center rounded-lg border border-black/10 px-4 text-sm font-bold transition hover:border-black/30"
            onClick={onCancel}
            type="button"
          >
            Cancelar
          </button>
          <button
            className="inline-flex h-10 items-center justify-center gap-2 rounded-lg bg-[#0f6b4b] px-4 text-sm font-bold text-white transition hover:bg-[#0b553b]"
            onClick={onValidate}
            type="button"
          >
            <CheckCircle2 className="h-4 w-4" />
            Validar y crear cuadros
          </button>
        </div>
      </div>

      <div className="mt-4 flex flex-wrap gap-2">
        {categories.map((category) => (
          <button
            className={`rounded-lg border px-3 py-2 text-sm font-bold transition ${
              category.id === selectedId
                ? "border-[#c59b45] bg-[#fff3d6] text-[#111816]"
                : "border-black/10 bg-[#f7f7f3] text-[#111816] hover:border-[#c59b45]"
            }`}
            key={category.id}
            onClick={() => onSelect(category.id)}
            type="button"
          >
            {category.name}
            <span className="ml-2 text-xs opacity-70">
              {category.pairs.length}
            </span>
          </button>
        ))}
      </div>

      {activeCategory?.warnings.length ? (
        <div className="mt-4 rounded-lg border border-amber-300 bg-amber-50 p-3 text-sm font-semibold text-amber-900">
          {activeCategory.warnings.slice(0, 4).map((warning) => (
            <p key={warning}>{warning}</p>
          ))}
        </div>
      ) : null}

      <div className="mt-4 overflow-x-auto">
        <table className="w-full min-w-[780px] border-collapse text-left text-sm">
          <thead>
            <tr className="bg-[#07110d] text-white">
              <th className="px-3 py-3 font-bold">#</th>
              <th className="px-3 py-3 font-bold">Columna B</th>
              <th className="px-3 py-3 font-bold">Columna C</th>
              <th className="px-3 py-3 font-bold">Restriccion D</th>
              <th className="px-3 py-3 font-bold">Comentario E</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((pair) => (
              <tr className="border-b border-black/10" key={pair.id}>
                <td className="px-3 py-3 font-bold">{pair.seed}</td>
                <td className="px-3 py-3">{pair.playerOne}</td>
                <td className="px-3 py-3">{pair.playerTwo}</td>
                <td className="px-3 py-3 text-[#0f6b4b]">
                  {pair.restriction || "-"}
                </td>
                <td className="px-3 py-3 text-black/60">
                  {pair.comment || "-"}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {(activeCategory?.pairs.length ?? 0) > previewLimit ? (
        <p className="mt-3 text-sm font-semibold text-black/60">
          Mostrando {previewLimit} de {activeCategory?.pairs.length} parejas.
        </p>
      ) : null}
    </section>
  );
}

function BracketBoard({
  categoryId,
  draw,
  emptyText,
  onPick,
  readOnly = false,
  schedule,
  subtitle,
}: {
  categoryId: string;
  draw: Draw;
  emptyText?: string;
  onPick?: (match: Match, team: Team) => void;
  readOnly?: boolean;
  schedule: Record<string, ScheduleAssignment>;
  subtitle: string;
}) {
  const isMain = draw.accent === "main";
  const scrollRef = useRef<HTMLDivElement | null>(null);
  const boardId = `bracket-${draw.accent}-${domSlug(draw.title)}`;

  function scrollToRound(roundIndex: number) {
    const container = scrollRef.current;
    const target = document.getElementById(`${boardId}-${roundIndex}`);

    if (!container || !target) return;

    const scrollLeft =
      target.getBoundingClientRect().left -
      container.getBoundingClientRect().left +
      container.scrollLeft;

    container.scrollTo({
      behavior: "smooth",
      left: scrollLeft,
    });
  }

  return (
    <section className="min-w-0 rounded-xl border border-black/10 bg-white p-4 shadow-sm">
      <div className="mb-4 flex flex-col gap-2 sm:flex-row sm:items-end sm:justify-between">
        <div>
          <p
            className={`text-xs font-bold uppercase tracking-[0.22em] ${
              isMain ? "text-[#0f6b4b]" : "text-[#9b7732]"
            }`}
          >
            {isMain ? "Principal" : "Segunda oportunidad"}
          </p>
          <h2 className="mt-1 text-2xl font-black">{draw.title}</h2>
          <p className="mt-1 text-sm font-semibold text-black/55">{subtitle}</p>
        </div>
      </div>

      {!draw.rounds.length ? (
        <div className="rounded-lg border border-dashed border-black/20 bg-[#f7f7f3] p-8 text-center text-sm font-bold text-black/50">
          {emptyText || "Todavia no hay suficientes parejas."}
        </div>
      ) : (
        <>
          <div className="mb-3 flex gap-2 overflow-x-auto pb-1 md:hidden">
            {draw.rounds.map((round, roundIndex) => (
              <button
                className={`shrink-0 rounded-lg px-3 py-2 text-sm font-black ${
                  isMain
                    ? "bg-[#e8f3ee] text-[#0b553b]"
                    : "bg-[#fff3d6] text-[#7a5818]"
                }`}
                key={round.name}
                onClick={() => scrollToRound(roundIndex)}
                type="button"
              >
                {round.name}
              </button>
            ))}
          </div>

          <div
            className="max-h-[72vh] touch-pan-x overflow-auto overscroll-contain rounded-xl bg-[#f7f7f3] p-2 pb-3 md:max-h-none md:overflow-x-auto md:rounded-none md:bg-transparent md:p-0 md:pb-2"
            ref={scrollRef}
          >
          <div className="flex min-w-max snap-x snap-mandatory gap-4 pr-8 md:snap-none md:gap-10">
            {draw.rounds.map((round, roundIndex) => (
              <div
                className="flex w-[300px] shrink-0 snap-start flex-col gap-3 md:w-[280px]"
                id={`${boardId}-${roundIndex}`}
                key={round.name}
              >
                <div
                  className={`rounded-lg px-3 py-2 text-sm font-black ${
                    isMain
                      ? "bg-[#e8f3ee] text-[#0b553b]"
                      : "bg-[#fff3d6] text-[#7a5818]"
                  }`}
                >
                  {round.name}
                </div>

                <div
                  className="flex flex-col gap-4 pt-0 md:gap-[var(--round-gap)] md:pt-[var(--round-padding)]"
                  style={bracketRoundStyle(roundIndex)}
                >
                  {round.matches.map((match, matchIndex) => (
                    <MatchCard
                      accent={draw.accent}
                      hasSourceConnector={roundIndex < draw.rounds.length - 1}
                      hasTargetConnector={roundIndex > 0}
                      key={match.id}
                      match={match}
                      onPick={onPick}
                      readOnly={readOnly}
                      schedule={schedule[matchScheduleKey(categoryId, match.id)]}
                      sourceSide={matchIndex % 2 === 0 ? "top" : "bottom"}
                    />
                  ))}
                </div>
              </div>
            ))}
          </div>
          </div>
        </>
      )}
    </section>
  );
}

function MatchCard({
  accent,
  hasSourceConnector,
  hasTargetConnector,
  match,
  onPick,
  readOnly,
  schedule,
  sourceSide,
}: {
  accent: Draw["accent"];
  hasSourceConnector: boolean;
  hasTargetConnector: boolean;
  match: Match;
  onPick?: (match: Match, team: Team) => void;
  readOnly: boolean;
  schedule?: ScheduleAssignment;
  sourceSide: "top" | "bottom";
}) {
  const color =
    accent === "main"
      ? "border-[#0f6b4b] bg-[#0f6b4b] text-white"
      : "border-[#c59b45] bg-[#c59b45] text-black";

  return (
    <article
      className={`bracket-match relative rounded-xl border border-black/10 bg-[#fbfbf8] p-2 shadow-sm ${
        hasTargetConnector ? "bracket-target" : ""
      } ${
        hasSourceConnector ? `bracket-source bracket-source-${sourceSide}` : ""
      }`}
      style={
        {
          "--match-height": `${bracketMatchHeight}px`,
          height: `${bracketMatchHeight}px`,
        } as CSSProperties
      }
    >
      {hasTargetConnector ? (
        <span
          aria-hidden="true"
          className="absolute left-[-20px] top-1/2 z-0 hidden w-5 border-t-2 border-[#0f6b4b]/30 md:block"
        />
      ) : null}
      {hasSourceConnector ? (
        <span
          aria-hidden="true"
          className="absolute right-[-20px] top-1/2 z-0 hidden w-5 border-t-2 border-[#0f6b4b]/30 md:block"
        />
      ) : null}
      {hasSourceConnector && sourceSide === "top" ? (
        <span
          aria-hidden="true"
          className="absolute right-[-20px] top-1/2 z-0 hidden border-l-2 border-[#0f6b4b]/30 md:block"
          style={{ height: "calc(var(--round-gap) + var(--match-height))" }}
        />
      ) : null}
      <div className="mb-2 flex items-center justify-between gap-2 px-1">
        <span className={`rounded-md border px-2 py-1 text-xs font-black ${color}`}>
          {match.id}
        </span>
        <span className="truncate text-xs font-bold uppercase tracking-[0.14em] text-black/45">
          {match.label}
        </span>
      </div>

      <div
        className={`mb-2 flex h-7 items-center gap-2 rounded-lg px-2 text-xs font-black ${
          schedule?.conflict
            ? "bg-amber-100 text-amber-900"
            : "bg-black/[0.04] text-black/60"
        }`}
      >
        <CalendarClock className="h-3.5 w-3.5 shrink-0" />
        <span className="truncate">{scheduleLabel(schedule)}</span>
      </div>

      <div className="grid gap-2">
        <TeamButton
          match={match}
          onPick={onPick}
          readOnly={readOnly}
          side="A"
          team={match.sideA}
        />
        <TeamButton
          match={match}
          onPick={onPick}
          readOnly={readOnly}
          side="B"
          team={match.sideB}
        />
      </div>
    </article>
  );
}

function TeamButton({
  match,
  onPick,
  readOnly,
  side,
  team,
}: {
  match: Match;
  onPick?: (match: Match, team: Team) => void;
  readOnly: boolean;
  side: "A" | "B";
  team: Team | null;
}) {
  const other = side === "A" ? match.sideB : match.sideA;
  const canPick = !readOnly && Boolean(onPick) && isPlayable(team) && isPlayable(other);
  const isWinner = team && match.winner?.id === team.id;
  const isPlaceholder = !team || team.isPlaceholder;

  return (
    <button
      className={`h-[54px] overflow-hidden rounded-lg border px-3 py-2 text-left transition ${
        isWinner
          ? "border-[#0f6b4b] bg-[#dff2e9] text-[#073d2a]"
          : isPlaceholder
            ? "border-black/5 bg-black/[0.03] text-black/35"
            : "border-black/10 bg-white text-[#111816] hover:border-[#0f6b4b] hover:bg-[#f0faf5]"
      } ${canPick ? "cursor-pointer" : "cursor-default"}`}
      disabled={!canPick}
      onClick={() => team && onPick?.(match, team)}
      type="button"
    >
      <span className="flex items-start gap-2">
        {team?.seed ? (
          <span className="mt-0.5 rounded bg-black/10 px-1.5 py-0.5 text-[11px] font-black">
            {team.seed}
          </span>
        ) : null}
        <span className="min-w-0 flex-1">
          <span className="line-clamp-2 text-sm font-black leading-5">
            {team?.name || "BYE"}
          </span>
        </span>
      </span>
    </button>
  );
}

function StatCard({
  icon: Icon,
  label,
  value,
}: {
  icon: ComponentType<{ className?: string }>;
  label: string;
  value: number | string;
}) {
  return (
    <article className="rounded-xl border border-black/10 bg-white p-4 shadow-sm">
      <div className="flex items-center justify-between gap-3">
        <div>
          <p className="text-xs font-bold uppercase tracking-[0.18em] text-black/45">
            {label}
          </p>
          <p className="mt-1 text-3xl font-black">{value}</p>
        </div>
        <div className="flex h-11 w-11 items-center justify-center rounded-lg bg-[#e8f3ee] text-[#0f6b4b]">
          <Icon className="h-5 w-5" />
        </div>
      </div>
    </article>
  );
}

function Notice({
  icon: Icon,
  text,
  title,
  tone,
}: {
  icon: ComponentType<{ className?: string }>;
  text: string;
  title: string;
  tone: "warning";
}) {
  return (
    <section
      className={`flex items-start gap-3 rounded-xl border p-4 ${
        tone === "warning"
          ? "border-amber-300 bg-amber-50 text-amber-950"
          : ""
      }`}
    >
      <Icon className="mt-0.5 h-5 w-5 shrink-0" />
      <div>
        <h2 className="font-black">{title}</h2>
        <p className="text-sm font-semibold">{text}</p>
      </div>
    </section>
  );
}

function EmptyState() {
  return (
    <section className="rounded-xl border border-dashed border-black/20 bg-white p-8 text-center shadow-sm">
      <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-xl bg-[#e8f3ee] text-[#0f6b4b]">
        <Upload className="h-7 w-7" />
      </div>
      <h2 className="mt-4 text-2xl font-black">Sube el Excel del torneo</h2>
      <p className="mx-auto mt-2 max-w-xl text-sm font-semibold leading-6 text-black/55">
        La app mostrara una previsualizacion antes de crear los cuadros.
      </p>
    </section>
  );
}

function formatPublishedDate(value: string) {
  try {
    return new Intl.DateTimeFormat("es-ES", {
      dateStyle: "short",
      timeStyle: "short",
    }).format(new Date(value));
  } catch {
    return "";
  }
}

export function TournamentPublicPage({ slug }: { slug: string }) {
  const storageKey = publishedTournamentStorageKey(slug);
  const tournament = useSyncExternalStore(
    (onStoreChange) => subscribeStorageKey(storageKey, onStoreChange),
    () => readPublishedTournament(slug),
    () => null,
  );
  const [selectedCategoryId, setSelectedCategoryId] = useState("");

  const drawSets = useMemo(
    () =>
      buildDrawSets(
        tournament?.categories ?? [],
        tournament?.mainSelectionsByCategory ?? {},
        tournament?.consolationSelectionsByCategory ?? {},
      ),
    [tournament],
  );
  const activeCategory =
    tournament?.categories.find(
      (category) =>
        category.id === (selectedCategoryId || tournament.selectedCategoryId),
    ) || tournament?.categories[0];
  const activeDrawSet =
    drawSets.find((drawSet) => drawSet.categoryId === activeCategory?.id) ||
    drawSets[0];
  const mainDraw = activeDrawSet?.mainDraw ?? emptyMainDraw;
  const consolationDraw = activeDrawSet?.consolationDraw ?? emptyConsolationDraw;
  const schedule = useMemo(
    () => buildGlobalSchedule(drawSets, tournament?.manualScheduleOverrides ?? {}),
    [drawSets, tournament?.manualScheduleOverrides],
  );

  if (!tournament) {
    return (
      <main className="min-h-screen bg-[#f5f4ef] text-[#111816]">
        <div className="mx-auto flex w-full max-w-[1100px] flex-col gap-6 px-4 py-8 sm:px-6 lg:px-8">
          <section className="rounded-xl border border-black/10 bg-white p-8 text-center shadow-sm">
            <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-xl bg-[#fff3d6] text-[#9b7732]">
              <AlertTriangle className="h-7 w-7" />
            </div>
            <h1 className="mt-4 text-3xl font-black">Cuadro no publicado</h1>
            <p className="mx-auto mt-2 max-w-xl text-sm font-semibold leading-6 text-black/55">
              Todavia no hay datos para este enlace en este entorno de pruebas.
              Cuando conectemos la base de datos, este enlace funcionara para
              cualquier visitante.
            </p>
          </section>
        </div>
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-[#f5f4ef] text-[#111816]">
      <div className="mx-auto flex w-full max-w-[1500px] flex-col gap-6 px-4 py-5 sm:px-6 lg:px-8">
        <header className="rounded-xl bg-[#07110d] p-5 text-white shadow-xl shadow-black/10 sm:p-7">
          <div className="flex flex-col gap-5 lg:flex-row lg:items-end lg:justify-between">
            <div>
              <p className="text-xs font-bold uppercase tracking-[0.24em] text-[#c59b45]">
                Cuadro publico
              </p>
              <h1 className="mt-2 text-3xl font-black tracking-normal sm:text-5xl">
                {tournament.title}
              </h1>
            </div>
            <div className="rounded-lg border border-white/15 px-4 py-3 text-sm font-bold text-white/75">
              Actualizado {formatPublishedDate(tournament.updatedAt)}
            </div>
          </div>
        </header>

        <section className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
          <StatCard
            icon={Users}
            label="Categorias"
            value={tournament.categories.length}
          />
          <StatCard
            icon={Trophy}
            label="Parejas"
            value={activeCategory?.pairs.length ?? 0}
          />
          <StatCard
            icon={CalendarClock}
            label="Sabado total"
            value={`${schedule.saturdayCount}/${schedule.total}`}
          />
          <StatCard icon={Medal} label="Estado" value="En juego" />
        </section>

        <section className="rounded-xl border border-black/10 bg-white p-4 shadow-sm">
          <div className="flex flex-wrap gap-2">
            {tournament.categories.map((category) => (
              <button
                className={`rounded-lg border px-3 py-2 text-sm font-bold transition ${
                  category.id === activeCategory?.id
                    ? "border-[#0f6b4b] bg-[#0f6b4b] text-white"
                    : "border-black/10 bg-[#f7f7f3] text-[#111816] hover:border-[#0f6b4b]"
                }`}
                key={category.id}
                onClick={() => setSelectedCategoryId(category.id)}
                type="button"
              >
                {category.name}
                <span className="ml-2 text-xs opacity-70">
                  {category.pairs.length}
                </span>
              </button>
            ))}
          </div>
        </section>

        <div className="grid gap-6">
          <BracketBoard
            categoryId={activeCategory?.id || ""}
            draw={mainDraw}
            readOnly
            schedule={schedule.assignments}
            subtitle="Resultados y horarios oficiales del cuadro principal."
          />
          <BracketBoard
            categoryId={activeCategory?.id || ""}
            draw={consolationDraw}
            emptyText="La consolacion se llenara cuando avance el torneo."
            readOnly
            schedule={schedule.assignments}
            subtitle="Cuadro de consolacion actualizado desde el panel admin."
          />
        </div>
      </div>
    </main>
  );
}
