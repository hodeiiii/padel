"use client";

import {
  AlertTriangle,
  CalendarClock,
  CheckCircle2,
  FileSpreadsheet,
  Medal,
  RotateCcw,
  Trophy,
  Upload,
  Users,
} from "lucide-react";
import {
  useMemo,
  useRef,
  useState,
  type ComponentType,
  type CSSProperties,
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

const previewLimit = 24;
const courtCount = 8;
const bracketMatchHeight = 210;
const bracketBaseGap = 16;
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

function buildGlobalSchedule(drawSets: CategoryDrawSet[]): ScheduleResult {
  const groups = globalChronologicalMatchGroups(drawSets);
  const slots = tournamentSlots();
  const timeGroups = scheduleTimeGroups(slots);
  const assignments: Record<string, ScheduleAssignment> = {};
  const usedSlotIndexes = new Set<number>();
  let latestGroup = timeGroups.length - 1;
  let conflicts = 0;

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
            conflicts += 1;
            return;
          }

          assignments[matchScheduleKey(item.categoryId, item.match.id)] = {
            ...fallbackSlot,
            conflict: !compatibleSlot,
          };

          if (!compatibleSlot) conflicts += 1;

          usedSlotIndexes.add(fallbackSlot.slotIndex);
          earliestUsedGroup = Math.min(earliestUsedGroup, timeGroup.groupIndex);
        });

      latestGroup = earliestUsedGroup - 1;
    });

  return {
    assignments,
    conflicts,
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
  const [previewCategories, setPreviewCategories] = useState<CategoryData[]>([]);
  const [previewCategoryId, setPreviewCategoryId] = useState("");
  const [categories, setCategories] = useState<CategoryData[]>([]);
  const [selectedCategoryId, setSelectedCategoryId] = useState("");
  const [mainSelectionsByCategory, setMainSelectionsByCategory] =
    useState<CategorySelectionMap>({});
  const [consolationSelectionsByCategory, setConsolationSelectionsByCategory] =
    useState<CategorySelectionMap>({});
  const [importError, setImportError] = useState("");

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
      categories.map((category) => {
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
      }),
    [categories, consolationSelectionsByCategory, mainSelectionsByCategory],
  );

  const activeDrawSet =
    drawSets.find((drawSet) => drawSet.categoryId === activeCategory?.id) ||
    drawSets[0];
  const mainDraw = activeDrawSet?.mainDraw ?? emptyMainDraw;
  const consolationDraw = activeDrawSet?.consolationDraw ?? emptyConsolationDraw;
  const schedule = useMemo(() => buildGlobalSchedule(drawSets), [drawSets]);

  const firstRoundResolved = useMemo(() => {
    const matches = mainDraw.rounds[0]?.matches ?? [];

    return matches.filter((match) => match.loser).length;
  }, [mainDraw]);

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
  }

  function selectCategory(id: string) {
    setSelectedCategoryId(id);
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

  function resetAll() {
    setPreviewCategories([]);
    setPreviewCategoryId("");
    setCategories([]);
    setSelectedCategoryId("");
    setMainSelectionsByCategory({});
    setConsolationSelectionsByCategory({});
    setImportError("");
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
        ) : null}

        {categories.length && schedule.conflicts ? (
          <Notice
            icon={AlertTriangle}
            tone="warning"
            title="Revisar horarios"
            text={`${schedule.conflicts} partido(s) no encajan perfectamente con las restricciones y se han marcado en amarillo.`}
          />
        ) : null}

        {categories.length ? (
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
        ) : (
          <EmptyState />
        )}
      </div>
    </main>
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
  schedule,
  subtitle,
}: {
  categoryId: string;
  draw: Draw;
  emptyText?: string;
  onPick: (match: Match, team: Team) => void;
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
  schedule,
  sourceSide,
}: {
  accent: Draw["accent"];
  hasSourceConnector: boolean;
  hasTargetConnector: boolean;
  match: Match;
  onPick: (match: Match, team: Team) => void;
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
        <TeamButton match={match} onPick={onPick} side="A" team={match.sideA} />
        <TeamButton match={match} onPick={onPick} side="B" team={match.sideB} />
      </div>
    </article>
  );
}

function TeamButton({
  match,
  onPick,
  side,
  team,
}: {
  match: Match;
  onPick: (match: Match, team: Team) => void;
  side: "A" | "B";
  team: Team | null;
}) {
  const other = side === "A" ? match.sideB : match.sideA;
  const canPick = isPlayable(team) && isPlayable(other);
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
      onClick={() => team && onPick(match, team)}
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
