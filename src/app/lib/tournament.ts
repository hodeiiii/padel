export type SheetCell = string | number | boolean | null | undefined;

export type DayKey = "jueves" | "viernes" | "sabado";
export type RuleMode = "available" | "blocked";
export type DrawKind = "main" | "consolation";

export type ScheduleWindow = {
  from: string;
  to: string;
};

export type TournamentScheduleConfig = {
  dayPriority: DayKey[];
  dayWindows: Record<DayKey, ScheduleWindow[]>;
};

export type RestrictionRule = {
  id: string;
  mode: RuleMode;
  day: DayKey;
  from: string;
  to: string;
  source: string;
  confidence: number;
};

export type Pair = {
  id: string;
  seed: number;
  playerOne: string;
  playerTwo: string;
  level: string;
  rawRestriction: string;
  rawNotes: string;
  rules: RestrictionRule[];
  review: boolean;
  reviewReasons: string[];
};

export type CategoryData = {
  id: string;
  name: string;
  sourceRow: number;
  pairs: Pair[];
  warnings: string[];
};

export type Team = {
  id: string;
  pairId?: string;
  seed?: number;
  name: string;
  fullName?: string;
  possibleTeams?: Team[];
  rules?: RestrictionRule[];
  isPlaceholder?: boolean;
};

export type Match = {
  id: string;
  label: string;
  roundIndex: number;
  matchIndex: number;
  draw: DrawKind;
  sideA: Team | null;
  sideB: Team | null;
  winner: Team | null;
  loser: Team | null;
  possibleTeams: Team[];
};

export type Round = {
  id: string;
  name: string;
  matches: Match[];
};

export type Draw = {
  kind: DrawKind;
  rounds: Round[];
};

export type CategoryDrawSet = {
  categoryId: string;
  categoryName: string;
  mainDraw: Draw;
  consolationDraw: Draw;
};

export type SelectionMap = Record<string, string>;
export type SelectionByCategory = Record<string, SelectionMap>;

export type ManualScheduleOverride = {
  court: number;
  day: DayKey;
  time: string;
};

export type ManualScheduleMap = Record<string, ManualScheduleOverride>;
export type PairLockMap = Record<string, string>;
export type ManualPairLockMap = Record<string, PairLockMap>;

export type ScheduleSlot = {
  court: number;
  day: DayKey;
  dayLabel: string;
  time: string;
  minutes: number;
  slotIndex: number;
};

export type ScheduleAssignment = ScheduleSlot & {
  conflict?: boolean;
  manual?: boolean;
};

export type ScheduleResult = {
  assignments: Record<string, ScheduleAssignment>;
  conflicts: number;
  saturdayCount: number;
  total: number;
};

export type ScheduleSummaryRow = {
  categoryId: string;
  categoryName: string;
  court: number | null;
  day: DayKey | null;
  dayLabel: string;
  drawName: string;
  match: string;
  matchId: string;
  manual: boolean;
  conflict: boolean;
  minutes: number;
  roundName: string;
  scheduleKey: string;
  time: string;
};

export type AdminState = {
  activeCategoryId: string;
  categories: CategoryData[];
  consolationSelectionsByCategory: SelectionByCategory;
  mainSelectionsByCategory: SelectionByCategory;
  manualPairLocks: ManualPairLockMap;
  manualScheduleOverrides: ManualScheduleMap;
  publishedSlug: string;
  scheduleConfig: TournamentScheduleConfig;
};

export type PublishedTournament = {
  categories: CategoryData[];
  consolationSelectionsByCategory: SelectionByCategory;
  mainSelectionsByCategory: SelectionByCategory;
  manualScheduleOverrides: ManualScheduleMap;
  name: string;
  publishedAt: string;
  scheduleConfig: TournamentScheduleConfig;
  slug: string;
};

export const adminStorageKey = "padel-admin-state-v2";
export const adminAuthStorageKey = "padel-admin-auth-v1";
export const publishedStoragePrefix = "padel-public-v2:";
export const defaultExcelPath = "/CUADROS%20PARA%20IA%202026.xlsx";
export const courtCount = 8;
export const matchDurationMinutes = 60;

export const dayOptions: { key: DayKey; label: string }[] = [
  { key: "jueves", label: "Jueves" },
  { key: "viernes", label: "Viernes" },
  { key: "sabado", label: "Sabado" },
];

const dayOrder: Record<DayKey, number> = {
  jueves: 0,
  viernes: 1,
  sabado: 2,
};

export const defaultScheduleConfig: TournamentScheduleConfig = {
  dayPriority: ["sabado", "viernes", "jueves"],
  dayWindows: {
    jueves: [
      { from: "10:30", to: "13:30" },
      { from: "17:00", to: "22:30" },
    ],
    viernes: [
      { from: "10:30", to: "13:30" },
      { from: "17:00", to: "22:30" },
    ],
    sabado: [{ from: "08:30", to: "21:30" }],
  },
};

export const emptyAdminState: AdminState = {
  activeCategoryId: "",
  categories: [],
  consolationSelectionsByCategory: {},
  mainSelectionsByCategory: {},
  manualPairLocks: {},
  manualScheduleOverrides: {},
  publishedSlug: "",
  scheduleConfig: defaultScheduleConfig,
};

export function dayLabel(day: DayKey) {
  return dayOptions.find((option) => option.key === day)?.label ?? day;
}

export function minutesFromTime(time: string) {
  const [hoursText, minutesText] = time.split(":");
  const hours = Number.parseInt(hoursText || "0", 10);
  const minutes = Number.parseInt(minutesText || "0", 10);

  return hours * 60 + minutes;
}

export function formatMinutes(minutes: number) {
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;

  return `${hours.toString().padStart(2, "0")}:${mins
    .toString()
    .padStart(2, "0")}`;
}

function normalizeTimeValue(value: string | null | undefined) {
  const match = /^(\d{1,2}):(\d{2})$/.exec(String(value ?? "").trim());

  if (!match) return "";

  const hours = Number.parseInt(match[1], 10);
  const minutes = Number.parseInt(match[2], 10);

  if (
    Number.isNaN(hours) ||
    Number.isNaN(minutes) ||
    hours < 0 ||
    hours > 23 ||
    minutes < 0 ||
    minutes > 59
  ) {
    return "";
  }

  return formatMinutes(hours * 60 + minutes);
}

export function normalizeScheduleConfig(
  scheduleConfig?: TournamentScheduleConfig,
): TournamentScheduleConfig {
  const source = scheduleConfig ?? defaultScheduleConfig;
  const dayWindows = dayOptions.reduce(
    (acc, day) => {
      const fallbackWindows = defaultScheduleConfig.dayWindows[day.key];
      const configuredWindows = source.dayWindows?.[day.key];
      const sourceWindows =
        Array.isArray(configuredWindows) && configuredWindows.length
          ? configuredWindows
          : fallbackWindows;

      acc[day.key] = sourceWindows.map((window, index) => {
        const fallback = fallbackWindows[index] ?? fallbackWindows[0];

        return {
          from: normalizeTimeValue(window.from) || fallback.from,
          to: normalizeTimeValue(window.to) || fallback.to,
        };
      });

      return acc;
    },
    {} as Record<DayKey, ScheduleWindow[]>,
  );
  const dayPriority: DayKey[] = [];
  const configuredPriority = Array.isArray(source.dayPriority)
    ? source.dayPriority
    : [];

  [...configuredPriority, ...defaultScheduleConfig.dayPriority].forEach(
    (day) => {
      const valid = dayOptions.some((option) => option.key === day);

      if (valid && !dayPriority.includes(day)) dayPriority.push(day);
    },
  );

  return {
    dayPriority,
    dayWindows,
  };
}

function scheduleWindowsForDay(
  day: DayKey,
  scheduleConfig: TournamentScheduleConfig = defaultScheduleConfig,
) {
  return normalizeScheduleConfig(scheduleConfig).dayWindows[day]
    .map((window) => ({
      from: minutesFromTime(window.from),
      to: minutesFromTime(window.to),
    }))
    .filter((window) => window.to > window.from);
}

export function timeOptionsForDay(
  day: DayKey,
  scheduleConfig: TournamentScheduleConfig = defaultScheduleConfig,
) {
  return scheduleWindowsForDay(day, scheduleConfig).flatMap((window) => {
    const times: string[] = [];

    for (
      let minutes = window.from;
      minutes + matchDurationMinutes <= window.to;
      minutes += matchDurationMinutes
    ) {
      times.push(formatMinutes(minutes));
    }

    return times;
  });
}

export function ruleModeLabel(mode: RuleMode) {
  return mode === "available" ? "Solo puede" : "No puede";
}

export function cleanCell(value: SheetCell) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

export function normalizeText(value: string) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/;/g, ":")
    .replace(/\s+/g, " ")
    .trim();
}

export function properCaseName(value: string) {
  return value
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map((part) =>
      part
        .split("-")
        .map((piece) =>
          piece
            ? `${piece.charAt(0).toLocaleUpperCase("es-ES")}${piece
                .slice(1)
                .toLocaleLowerCase("es-ES")}`
            : piece,
        )
        .join("-"),
    )
    .join(" ");
}

export function teamShortName(pair: Pick<Pair, "playerOne" | "playerTwo">) {
  return `${firstSurname(pair.playerOne)} / ${firstSurname(pair.playerTwo)}`;
}

export function pairFullName(pair: Pick<Pair, "playerOne" | "playerTwo">) {
  return `${properCaseName(pair.playerOne)} / ${properCaseName(pair.playerTwo)}`;
}

function firstSurname(player: string) {
  const particles = new Set(["de", "del", "la", "las", "los", "y"]);
  const tokens = properCaseName(player).split(/\s+/).filter(Boolean);

  if (tokens.length <= 1) return tokens[0] ?? player.trim();

  return tokens.slice(1).find((token) => !particles.has(token.toLowerCase())) ??
    tokens[1];
}

function makeRuleId(seed: string, index: number) {
  return `${seed}-rule-${index + 1}`;
}

function dayBounds(day: DayKey) {
  const windows = scheduleWindowsForDay(day);

  return {
    from: windows[0]?.from ?? 0,
    to: windows[windows.length - 1]?.to ?? 24 * 60,
  };
}

function clampToDay(day: DayKey, minutes: number) {
  const bounds = dayBounds(day);

  return Math.max(bounds.from, Math.min(bounds.to, minutes));
}

function parseRestrictionTime(
  hourText: string,
  minuteText?: string,
  suffix?: string,
) {
  let hours = Number.parseInt(hourText, 10);
  const minutes = Number.parseInt(minuteText || "0", 10);
  const normalizedSuffix = suffix?.toLowerCase();

  if (normalizedSuffix === "pm" && hours < 12) hours += 12;
  if (normalizedSuffix === "am" && hours === 12) hours = 0;

  return hours * 60 + minutes;
}

function restrictionText(value: string) {
  return normalizeText(value)
    .replace(/\bveirenes\b/g, "viernes")
    .replace(/\bvierens\b/g, "viernes")
    .replace(/\bjuever\b/g, "jueves")
    .replace(/\bjueves?\b/g, "jueves")
    .replace(/\bsabado\b/g, "sabado")
    .replace(/\bmananas\b/g, "manana")
    .replace(/\bmediodias\b/g, "mediodia")
    .replace(/\bdespues\b/g, "despues")
    .replace(/\balas\b/g, "a las")
    .replace(/(\d{1,2})\.(\d{2})/g, "$1:$2")
    .replace(/(\d{1,2}),(\d{2})/g, "$1:$2");
}

function splitRestrictionSegments(text: string) {
  let next = text;

  for (let index = 0; index < 6; index += 1) {
    const previous = next;
    next = next.replace(
      /((?:\+|-|=|\bno\b|\bentre\b|\bpartir\b|\bdesde\b)[^;,\n]*?)\s+y\s+(jueves|viernes|sabado)\b/g,
      "$1; $2",
    );

    if (next === previous) break;
  }

  return next
    .split(/[,\n;]+/)
    .map((segment) => segment.trim())
    .filter(Boolean);
}

function daysInText(text: string): DayKey[] {
  const days: DayKey[] = [];

  if (/\btodos los dias\b/.test(text) || /\btodo el dia\b/.test(text)) {
    return ["jueves", "viernes", "sabado"];
  }
  if (/\bjueves\b|\bj\b/.test(text)) days.push("jueves");
  if (/\bviernes\b|\bv\b/.test(text)) days.push("viernes");
  if (/\bsabado\b|\bs\b/.test(text)) days.push("sabado");

  return [...new Set(days)];
}

function addRule(
  rules: RestrictionRule[],
  seed: string,
  mode: RuleMode,
  day: DayKey,
  fromMinutes: number,
  toMinutes: number,
  source: string,
  confidence: number,
) {
  const from = clampToDay(day, fromMinutes);
  const to = clampToDay(day, toMinutes);

  if (to <= from) return;

  const duplicate = rules.some(
    (rule) =>
      rule.mode === mode &&
      rule.day === day &&
      rule.from === formatMinutes(from) &&
      rule.to === formatMinutes(to),
  );

  if (duplicate) return;

  rules.push({
    confidence,
    day,
    from: formatMinutes(from),
    id: makeRuleId(seed, rules.length),
    mode,
    source,
    to: formatMinutes(to),
  });
}

function parseRulesFromSegment(
  segment: string,
  seed: string,
  rules: RestrictionRule[],
) {
  const days = daysInText(segment);

  if (!days.length) return;

  const source = segment;
  const time = String.raw`(\d{1,2})(?::?(\d{2}))?\s*(am|pm|h)?`;
  const deniedInterval = new RegExp(
    String.raw`(?:no\s+puede|no\s+podemos|no)?\s*(?:entre)\s+(?:las\s+)?${time}\s+(?:y|a)\s+(?:las\s+)?${time}`,
    "g",
  );
  const availableRange = new RegExp(
    String.raw`\+\s*${time}\s*(?:-|hasta|a)\s*${time}`,
    "g",
  );
  const exactPattern = new RegExp(String.raw`=\s*${time}`, "g");
  const minimumPattern = new RegExp(
    String.raw`(?:\+|a partir de|desde|despues de)\s*(?:las\s+)?${time}`,
    "g",
  );
  const maximumPattern = new RegExp(
    String.raw`(?:-|antes de|hasta)\s*(?:las\s+)?${time}`,
    "g",
  );
  const blockedSpans: [number, number][] = [];

  for (const match of segment.matchAll(deniedInterval)) {
    blockedSpans.push([match.index ?? 0, (match.index ?? 0) + match[0].length]);
    days.forEach((day) =>
      addRule(
        rules,
        seed,
        "blocked",
        day,
        parseRestrictionTime(match[1], match[2], match[3]),
        parseRestrictionTime(match[4], match[5], match[6]),
        source,
        92,
      ),
    );
  }

  if (/\b(no|ni|sin)\b/.test(segment) && /\bmanana\b/.test(segment)) {
    days.forEach((day) =>
      addRule(
        rules,
        seed,
        "blocked",
        day,
        dayBounds(day).from,
        14 * 60,
        source,
        86,
      ),
    );
  }

  if (/\b(no|ni|sin)\b/.test(segment) && /\btarde\b/.test(segment)) {
    days.forEach((day) =>
      addRule(
        rules,
        seed,
        "blocked",
        day,
        14 * 60,
        dayBounds(day).to,
        source,
        86,
      ),
    );
  }

  if (/\bno\b/.test(segment) && !segment.includes("entre")) {
    const hasSpecificTime = /(?:\+|-|=|\d{1,2}:?\d{0,2})/.test(
      segment.replace(/\b(jueves|viernes|sabado)\b/g, ""),
    );

    if (!hasSpecificTime) {
      days.forEach((day) =>
        addRule(
          rules,
          seed,
          "blocked",
          day,
          dayBounds(day).from,
          dayBounds(day).to,
          source,
          92,
        ),
      );
    }
  }

  for (const match of segment.matchAll(availableRange)) {
    blockedSpans.push([match.index ?? 0, (match.index ?? 0) + match[0].length]);
    days.forEach((day) =>
      addRule(
        rules,
        seed,
        "available",
        day,
        parseRestrictionTime(match[1], match[2], match[3]),
        parseRestrictionTime(match[4], match[5], match[6]),
        source,
        95,
      ),
    );
  }

  const spanContains = (index: number) =>
    blockedSpans.some(([from, to]) => index >= from && index < to);

  for (const match of segment.matchAll(exactPattern)) {
    if (spanContains(match.index ?? 0)) continue;

    days.forEach((day) => {
      const startsAt = parseRestrictionTime(match[1], match[2], match[3]);

      addRule(
        rules,
        seed,
        "available",
        day,
        startsAt,
        startsAt + matchDurationMinutes,
        source,
        68,
      );
    });
  }

  for (const match of segment.matchAll(minimumPattern)) {
    if (spanContains(match.index ?? 0)) continue;

    days.forEach((day) =>
      addRule(
        rules,
        seed,
        "available",
        day,
        parseRestrictionTime(match[1], match[2], match[3]),
        dayBounds(day).to,
        source,
        88,
      ),
    );
  }

  for (const match of segment.matchAll(maximumPattern)) {
    if (spanContains(match.index ?? 0)) continue;

    days.forEach((day) =>
      addRule(
        rules,
        seed,
        "available",
        day,
        dayBounds(day).from,
        parseRestrictionTime(match[1], match[2], match[3]),
        source,
        84,
      ),
    );
  }
}

export function autocompleteRestrictionRules(
  compactRestriction: string,
  notes: string,
  seed: string,
) {
  const compact = restrictionText(compactRestriction);
  const natural = restrictionText(notes);
  const sourceText = compact || natural;
  const combined = `${compact} ${natural}`.trim();
  const rules: RestrictionRule[] = [];
  const reviewReasons: string[] = [];

  if (/\brevisar\b|\bmovil\b/.test(combined)) {
    reviewReasons.push("Marcado como revisar en el Excel");
  }

  if (
    natural &&
    /\b(a tope|cualquier hora|sin problema|libre)\b/.test(natural) &&
    (!compact || /\+\s*0?9(?::?00)?\s*am?/.test(compact))
  ) {
    return {
      review: false,
      reviewReasons,
      rules,
    };
  }

  splitRestrictionSegments(sourceText).forEach((segment) =>
    parseRulesFromSegment(segment, seed, rules),
  );

  if (!rules.length && combined && !/\b(a tope|cualquier hora|sin problema)\b/.test(combined)) {
    reviewReasons.push("No se ha podido convertir a reglas horarias");
  }

  if (/\bno podemos\b|\bno puede\b/.test(natural) && !rules.length) {
    reviewReasons.push("Hay texto negativo sin rango claro");
  }

  return {
    review: reviewReasons.length > 0,
    reviewReasons,
    rules,
  };
}

export function parseCategoriesFromRows(rows: SheetCell[][]) {
  const categories: CategoryData[] = [];
  let currentCategory: CategoryData | null = null;
  const nameCounts = new Map<string, number>();

  rows.forEach((row, rowIndex) => {
    const firstCell = row[0];
    const firstCellText = cleanCell(firstCell);
    const maybeCategory = cleanCell(row[1]);

    if (
      maybeCategory &&
      (!firstCellText || !Number.isFinite(Number(firstCellText))) &&
      normalizeText(maybeCategory).includes("categoria")
    ) {
      const baseName = maybeCategory.replace(/\s+/g, " ").trim();
      const count = (nameCounts.get(baseName) ?? 0) + 1;
      const displayName = count > 1 ? `${baseName} (${count})` : baseName;

      nameCounts.set(baseName, count);
      currentCategory = {
        id: `cat-${categories.length + 1}-${rowIndex + 1}`,
        name: displayName,
        pairs: [],
        sourceRow: rowIndex + 1,
        warnings: [],
      };
      categories.push(currentCategory);
      return;
    }

    if (!currentCategory) return;

    const seed = Number.parseInt(cleanCell(firstCell), 10);
    const playerOne = cleanCell(row[1]);
    const playerTwo = cleanCell(row[2]);

    if (!Number.isFinite(seed) || !playerOne || !playerTwo) return;

    const rawRestriction = cleanCell(row[4]);
    const rawNotes = cleanCell(row[5]);
    const pairSeed = `${currentCategory.id}-${seed}`;
    const suggestion = autocompleteRestrictionRules(
      rawRestriction,
      rawNotes,
      pairSeed,
    );

    currentCategory.pairs.push({
      id: `pair-${currentCategory.id}-${seed}`,
      level: cleanCell(row[3]),
      playerOne: properCaseName(playerOne),
      playerTwo: properCaseName(playerTwo),
      rawNotes,
      rawRestriction,
      review: suggestion.review,
      reviewReasons: suggestion.reviewReasons,
      rules: suggestion.rules,
      seed,
    });
  });

  categories.forEach((category) => {
    if (!category.pairs.length) {
      category.warnings.push("Categoria sin parejas detectadas");
    }

    if (category.pairs.length % 2 !== 0) {
      category.warnings.push("Numero impar de parejas: el cuadro tendra bye");
    }
  });

  return categories.filter((category) => category.pairs.length);
}

function nextPowerOfTwo(value: number) {
  let size = 1;

  while (size < value) size *= 2;

  return Math.max(size, 2);
}

export function teamFromPair(pair: Pair): Team {
  return {
    fullName: pairFullName(pair),
    id: pair.id,
    name: teamShortName(pair),
    pairId: pair.id,
    rules: pair.rules,
    seed: pair.seed,
  };
}

function uniquePlayableTeams(teams: Team[]) {
  const seen = new Set<string>();
  const unique: Team[] = [];

  teams.forEach((team) => {
    if (team.isPlaceholder || seen.has(team.id)) return;

    seen.add(team.id);
    unique.push(team);
  });

  return unique;
}

function placeholderTeam(name: string, possibleTeams: Team[] = []): Team {
  return {
    id: `placeholder-${name}`,
    isPlaceholder: true,
    name,
    possibleTeams: uniquePlayableTeams(possibleTeams),
  };
}

function isPlayableTeam(team: Team | null): team is Team {
  return Boolean(team && !team.isPlaceholder);
}

function schedulingTeamsForSide(team: Team | null) {
  if (!team) return [];
  if (!team.isPlaceholder) return [team];

  return team.possibleTeams ?? [];
}

function possibleTeamsFromSides(sideA: Team | null, sideB: Team | null) {
  return uniquePlayableTeams([
    ...schedulingTeamsForSide(sideA),
    ...schedulingTeamsForSide(sideB),
  ]);
}

function sideCanExist(team: Team | null) {
  return schedulingTeamsForSide(team).length > 0;
}

export function isPlayableMatch(match: Match) {
  return isPlayableTeam(match.sideA) && isPlayableTeam(match.sideB);
}

function selectedWinner(
  sideA: Team | null,
  sideB: Team | null,
  selectedId: string | undefined,
) {
  if (sideCanExist(sideA) && sideB === null) return sideA;
  if (sideA === null && sideCanExist(sideB)) return sideB;
  if (selectedId && sideA?.id === selectedId) return sideA;
  if (selectedId && sideB?.id === selectedId) return sideB;

  return null;
}

function selectedLoser(
  sideA: Team | null,
  sideB: Team | null,
  winner: Team | null,
) {
  if (!winner || !isPlayableTeam(sideA) || !isPlayableTeam(sideB)) return null;

  return winner.id === sideA.id ? sideB : sideA;
}

export function roundName(totalSize: number, roundIndex: number) {
  const remaining = totalSize / 2 ** roundIndex;

  if (remaining >= 32) return "Dieciseisavos";
  if (remaining === 16) return "Octavos";
  if (remaining === 8) return "Cuartos";
  if (remaining === 4) return "Semifinales";
  if (remaining === 2) return "Final";

  return `Ronda ${roundIndex + 1}`;
}

function buildBracketFromTeams(
  teams: (Team | null)[],
  selections: SelectionMap,
  kind: DrawKind,
) {
  const size = nextPowerOfTwo(teams.length);
  const slots = [...teams];

  while (slots.length < size) slots.push(null);

  const rounds: Round[] = [];
  let previousSides = slots;
  const totalRounds = Math.log2(size);

  for (let roundIndex = 0; roundIndex < totalRounds; roundIndex += 1) {
    const matches: Match[] = [];
    const roundSize = previousSides.length;

    for (let index = 0; index < roundSize; index += 2) {
      const matchIndex = index / 2;
      const idPrefix = kind === "main" ? "P" : "C";
      const id = `${idPrefix}${roundIndex + 1}-${matchIndex + 1}`;
      const sideA = previousSides[index] ?? null;
      const sideB = previousSides[index + 1] ?? null;
      const possibleTeams = possibleTeamsFromSides(sideA, sideB);
      const winner = selectedWinner(sideA, sideB, selections[id]);
      const loser = selectedLoser(sideA, sideB, winner);

      matches.push({
        draw: kind,
        id,
        label: `${roundName(size, roundIndex)} ${matchIndex + 1}`,
        loser,
        matchIndex,
        possibleTeams,
        roundIndex,
        sideA,
        sideB,
        winner,
      });
    }

    rounds.push({
      id: `${kind}-round-${roundIndex + 1}`,
      matches,
      name: roundName(size, roundIndex),
    });

    previousSides = matches.map((match) => {
      if (match.winner) return match.winner;
      if (!match.possibleTeams.length) return null;

      return placeholderTeam(`Ganador ${match.id}`, match.possibleTeams);
    });
  }

  return {
    kind,
    rounds,
  };
}

export function buildDrawSet(
  category: CategoryData,
  mainSelections: SelectionMap = {},
  consolationSelections: SelectionMap = {},
): CategoryDrawSet {
  const mainTeams = category.pairs.map(teamFromPair);
  const mainDraw = buildBracketFromTeams(mainTeams, mainSelections, "main");
  const firstRound = mainDraw.rounds[0]?.matches ?? [];
  const consolationTeams = firstRound
    .filter((match) => isPlayableTeam(match.sideA) && isPlayableTeam(match.sideB))
    .map(
      (match) =>
        match.loser ?? placeholderTeam(`Perdedor ${match.id}`, match.possibleTeams),
    );
  const consolationDraw = consolationTeams.length
    ? buildBracketFromTeams(
        consolationTeams,
        consolationSelections,
        "consolation",
      )
    : { kind: "consolation" as const, rounds: [] };

  return {
    categoryId: category.id,
    categoryName: category.name,
    consolationDraw,
    mainDraw,
  };
}

export function buildDrawSets(
  categories: CategoryData[],
  mainSelectionsByCategory: SelectionByCategory,
  consolationSelectionsByCategory: SelectionByCategory,
) {
  return categories.map((category) =>
    buildDrawSet(
      category,
      mainSelectionsByCategory[category.id] ?? {},
      consolationSelectionsByCategory[category.id] ?? {},
    ),
  );
}

export function matchScheduleKey(categoryId: string, match: Match) {
  return `${categoryId}::${match.draw}::${match.id}`;
}

export function matchTeamsLabel(match: Match) {
  return `${match.sideA?.name ?? "Pendiente"} vs ${
    match.sideB?.name ?? "Pendiente"
  }`;
}

function makeSlotsForRange(
  day: DayKey,
  fromMinutes: number,
  toMinutes: number,
) {
  const slots: ScheduleSlot[] = [];

  for (
    let minutes = fromMinutes;
    minutes + matchDurationMinutes <= toMinutes;
    minutes += matchDurationMinutes
  ) {
    for (let court = 1; court <= courtCount; court += 1) {
      slots.push({
        court,
        day,
        dayLabel: dayLabel(day),
        minutes,
        slotIndex: -1,
        time: formatMinutes(minutes),
      });
    }
  }

  return slots;
}

export function tournamentSlots(
  scheduleConfig: TournamentScheduleConfig = defaultScheduleConfig,
) {
  const normalizedConfig = normalizeScheduleConfig(scheduleConfig);

  return normalizedConfig.dayPriority
    .slice()
    .reverse()
    .flatMap((day) =>
      scheduleWindowsForDay(day, normalizedConfig).flatMap((window) =>
        makeSlotsForRange(day, window.from, window.to),
      ),
    )
    .map((slot, slotIndex) => ({ ...slot, slotIndex }));
}

function scheduleTimeGroups(slots: ScheduleSlot[]) {
  const groups: { groupIndex: number; slots: ScheduleSlot[] }[] = [];

  slots.forEach((slot) => {
    const current = groups.at(-1);

    if (
      current &&
      current.slots[0]?.day === slot.day &&
      current.slots[0]?.time === slot.time
    ) {
      current.slots.push(slot);
      return;
    }

    groups.push({
      groupIndex: groups.length,
      slots: [slot],
    });
  });

  return groups;
}

function rangeOverlaps(aFrom: number, aTo: number, bFrom: number, bTo: number) {
  return aFrom < bTo && bFrom < aTo;
}

export function teamFitsSlot(team: Team, slot: ScheduleSlot) {
  const rules = team.rules ?? [];
  const slotFrom = slot.minutes;
  const slotTo = slot.minutes + matchDurationMinutes;
  const dayRules = rules.filter((rule) => rule.day === slot.day);
  const blocked = dayRules.filter((rule) => rule.mode === "blocked");
  const available = dayRules.filter((rule) => rule.mode === "available");

  if (
    blocked.some((rule) =>
      rangeOverlaps(
        slotFrom,
        slotTo,
        minutesFromTime(rule.from),
        minutesFromTime(rule.to),
      ),
    )
  ) {
    return false;
  }

  if (!available.length) return true;

  return available.some(
    (rule) =>
      slotFrom >= minutesFromTime(rule.from) &&
      slotTo <= minutesFromTime(rule.to),
  );
}

export function matchFitsSlot(match: Match, slot: ScheduleSlot) {
  return possibleTeamsFromSides(match.sideA, match.sideB).every((team) =>
    teamFitsSlot(team, slot),
  );
}

export function matchNeedsSchedule(match: Match) {
  return sideCanExist(match.sideA) && sideCanExist(match.sideB);
}

function chronologicalMatchGroups(drawSet: CategoryDrawSet) {
  const groups: Match[][] = [];
  const maxRounds = Math.max(
    drawSet.mainDraw.rounds.length,
    drawSet.consolationDraw.rounds.length + 1,
  );

  for (let index = 0; index < maxRounds; index += 1) {
    const stageMatches = [
      ...(drawSet.mainDraw.rounds[index]?.matches ?? []),
      ...(index > 0
        ? drawSet.consolationDraw.rounds[index - 1]?.matches ?? []
        : []),
    ].filter(matchNeedsSchedule);

    if (stageMatches.length) groups.push(stageMatches);
  }

  return groups;
}

function globalChronologicalMatchGroups(drawSets: CategoryDrawSet[]) {
  const perCategory = drawSets.map((drawSet) => ({
    categoryId: drawSet.categoryId,
    categoryName: drawSet.categoryName,
    groups: chronologicalMatchGroups(drawSet),
  }));
  const totalStages = Math.max(
    0,
    ...perCategory.map((drawSet) => drawSet.groups.length),
  );
  const groups: { categoryId: string; categoryName: string; match: Match }[][] = [];

  for (let index = 0; index < totalStages; index += 1) {
    const stageMatches = perCategory.flatMap((drawSet) =>
      (drawSet.groups[index] ?? []).map((match) => ({
        categoryId: drawSet.categoryId,
        categoryName: drawSet.categoryName,
        match,
      })),
    );

    if (stageMatches.length) groups.push(stageMatches);
  }

  return groups;
}

export function manualOverrideToAssignment(
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

export function scheduleToManualOverride(
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

function scheduleSlotSignature(
  schedule: Pick<ScheduleAssignment, "court" | "day" | "time">,
) {
  return `${schedule.day}-${schedule.time}-${schedule.court}`;
}

export function buildGlobalSchedule(
  drawSets: CategoryDrawSet[],
  manualScheduleOverrides: ManualScheduleMap = {},
  scheduleConfig: TournamentScheduleConfig = defaultScheduleConfig,
): ScheduleResult {
  const groups = globalChronologicalMatchGroups(drawSets);
  const slots = tournamentSlots(scheduleConfig);
  const timeGroups = scheduleTimeGroups(slots);
  const slotBySignature = new Map(
    slots.map((slot) => [scheduleSlotSignature(slot), slot]),
  );
  const groupIndexBySlotIndex = new Map<number, number>();
  const assignments: Record<string, ScheduleAssignment> = {};
  const matchesByKey = new Map<string, Match>();
  const usedSlotIndexes = new Set<number>();
  const fixedGroupIndexes = new Map<string, number>();
  let latestGroup = timeGroups.length - 1;
  let unscheduledConflicts = 0;

  timeGroups.forEach((timeGroup) => {
    timeGroup.slots.forEach((slot) => {
      groupIndexBySlotIndex.set(slot.slotIndex, timeGroup.groupIndex);
    });
  });

  groups.forEach((group) => {
    group.forEach((item) => {
      matchesByKey.set(matchScheduleKey(item.categoryId, item.match), item.match);
    });
  });

  Object.entries(manualScheduleOverrides).forEach(([key, override]) => {
    if (!matchesByKey.has(key)) return;

    const manual = manualOverrideToAssignment(override);
    const slot = slotBySignature.get(scheduleSlotSignature(manual));

    assignments[key] = slot
      ? {
          ...slot,
          manual: true,
        }
      : manual;

    if (slot) {
      usedSlotIndexes.add(slot.slotIndex);
      fixedGroupIndexes.set(
        key,
        groupIndexBySlotIndex.get(slot.slotIndex) ?? latestGroup,
      );
    }
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
          const key = matchScheduleKey(item.categoryId, item.match);
          const fixedGroupIndex = fixedGroupIndexes.get(key);

          if (assignments[key]) {
            if (fixedGroupIndex !== undefined) {
              earliestUsedGroup = Math.min(earliestUsedGroup, fixedGroupIndex);
            }
            return;
          }

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
            timeGroup?.slots.find((slot) => !usedSlotIndexes.has(slot.slotIndex));

          if (!fallbackSlot || !timeGroup) {
            const overflowTimeGroup = timeGroups.findLast((candidateGroup) =>
              candidateGroup.slots.some(
                (slot) => !usedSlotIndexes.has(slot.slotIndex),
              ),
            );
            const overflowSlot = overflowTimeGroup?.slots.find(
              (slot) => !usedSlotIndexes.has(slot.slotIndex),
            );

            if (!overflowSlot || !overflowTimeGroup) {
              unscheduledConflicts += 1;
              return;
            }

            assignments[key] = {
              ...overflowSlot,
              conflict: true,
            };
            usedSlotIndexes.add(overflowSlot.slotIndex);
            earliestUsedGroup = Math.min(
              earliestUsedGroup,
              overflowTimeGroup.groupIndex,
            );
            return;
          }

          assignments[key] = {
            ...fallbackSlot,
            conflict: !compatibleSlot,
          };
          usedSlotIndexes.add(fallbackSlot.slotIndex);
          earliestUsedGroup = Math.min(earliestUsedGroup, timeGroup.groupIndex);
        });

      latestGroup = earliestUsedGroup - 1;
    });

  const conflictKeys = new Set<string>();
  const occupiedSlots = new Map<string, string[]>();

  Object.entries(assignments).forEach(([key, assignment]) => {
    const match = matchesByKey.get(key);
    const slotKey = scheduleSlotSignature(assignment);
    const slotMatches = occupiedSlots.get(slotKey) ?? [];

    slotMatches.push(key);
    occupiedSlots.set(slotKey, slotMatches);

    if (
      assignment.conflict ||
      assignment.slotIndex < 0 ||
      (match && !matchFitsSlot(match, assignment))
    ) {
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

export function buildScheduleSummaryRows(
  categories: CategoryData[],
  drawSets: CategoryDrawSet[],
  schedule: Record<string, ScheduleAssignment>,
): ScheduleSummaryRow[] {
  const categoryNames = new Map(categories.map((category) => [category.id, category.name]));
  const rows: ScheduleSummaryRow[] = [];

  drawSets.forEach((drawSet) => {
    [
      { draw: drawSet.mainDraw, drawName: "Principal" },
      { draw: drawSet.consolationDraw, drawName: "Consolacion" },
    ].forEach(({ draw, drawName }) => {
      draw.rounds.forEach((round) => {
        round.matches.filter(matchNeedsSchedule).forEach((match) => {
          const scheduleKey = matchScheduleKey(drawSet.categoryId, match);
          const assignment = schedule[scheduleKey];

          rows.push({
            categoryId: drawSet.categoryId,
            categoryName:
              categoryNames.get(drawSet.categoryId) ?? drawSet.categoryName,
            conflict: Boolean(assignment?.conflict),
            court: assignment?.court ?? null,
            day: assignment?.day ?? null,
            dayLabel: assignment?.dayLabel ?? "Sin dia",
            drawName,
            manual: Boolean(assignment?.manual),
            match: matchTeamsLabel(match),
            matchId: match.id,
            minutes: assignment?.minutes ?? Number.MAX_SAFE_INTEGER,
            roundName: round.name,
            scheduleKey,
            time: assignment?.time ?? "Sin hora",
          });
        });
      });
    });
  });

  return rows.sort((a, b) => {
    const dayA = a.day ? dayOrder[a.day] : 99;
    const dayB = b.day ? dayOrder[b.day] : 99;

    if (dayA !== dayB) return dayA - dayB;
    if (a.minutes !== b.minutes) return a.minutes - b.minutes;
    if ((a.court ?? 99) !== (b.court ?? 99)) return (a.court ?? 99) - (b.court ?? 99);

    return a.categoryName.localeCompare(b.categoryName);
  });
}

export function shufflePairs(pairs: Pair[]) {
  const next = [...pairs];

  for (let index = next.length - 1; index > 0; index -= 1) {
    const target = Math.floor(Math.random() * (index + 1));
    const current = next[index];

    next[index] = next[target];
    next[target] = current;
  }

  return next.map((pair, index) => ({
    ...pair,
    seed: index + 1,
  }));
}

export function shufflePairsKeepingLocks(
  pairs: Pair[],
  pairLocks: PairLockMap = {},
) {
  const lockedPairs = new Set<string>();
  const pairById = new Map(pairs.map((pair) => [pair.id, pair]));
  const next: (Pair | null)[] = Array.from({ length: pairs.length }, () => null);

  Object.entries(pairLocks).forEach(([slot, pairId]) => {
    const slotIndex = Number.parseInt(slot, 10);
    const pair = pairById.get(pairId);

    if (
      !Number.isInteger(slotIndex) ||
      slotIndex < 0 ||
      slotIndex >= pairs.length ||
      !pair ||
      lockedPairs.has(pairId)
    ) {
      return;
    }

    next[slotIndex] = pair;
    lockedPairs.add(pairId);
  });

  const shuffledUnlocked = shufflePairs(
    pairs.filter((pair) => !lockedPairs.has(pair.id)),
  );
  let unlockedIndex = 0;

  return next.map((pair, index) => ({
    ...(pair ?? shuffledUnlocked[unlockedIndex++]),
    seed: index + 1,
  }));
}

export function clearCategoryManualSchedules(
  manualScheduleOverrides: ManualScheduleMap,
  categoryId: string,
) {
  const prefix = `${categoryId}::`;
  const next = { ...manualScheduleOverrides };

  Object.keys(next).forEach((key) => {
    if (key.startsWith(prefix)) delete next[key];
  });

  return next;
}

export function slugFromName(value: string) {
  const slug = normalizeText(value)
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");

  return slug || "torneo-padel";
}

export function publicStorageKey(slug: string) {
  return `${publishedStoragePrefix}${slug}`;
}

export function scheduleLabel(schedule?: ScheduleAssignment) {
  if (!schedule) return "Sin horario";

  return `${schedule.dayLabel} ${schedule.time} - Pista ${schedule.court}`;
}

export function countReviewPairs(categories: CategoryData[]) {
  return categories.reduce(
    (total, category) => total + category.pairs.filter((pair) => pair.review).length,
    0,
  );
}

export function countPairs(categories: CategoryData[]) {
  return categories.reduce((total, category) => total + category.pairs.length, 0);
}
