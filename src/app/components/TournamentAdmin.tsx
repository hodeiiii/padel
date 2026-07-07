"use client";

import {
  AlertTriangle,
  CheckCircle2,
  Copy,
  Crown,
  ExternalLink,
  FileSpreadsheet,
  Globe2,
  LockKeyhole,
  LockOpen,
  Plus,
  RotateCcw,
  Search,
  Share2,
  Shuffle,
  Trash2,
  Trophy,
  Upload,
  Users,
  X,
} from "lucide-react";
import {
  useEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
  type ChangeEvent,
} from "react";
import * as XLSX from "xlsx";
import {
  adminAuthStorageKey,
  adminStorageKey,
  buildDrawSets,
  buildGlobalSchedule,
  buildScheduleSummaryRows,
  clearCategoryManualSchedules,
  countPairs,
  countReviewPairs,
  courtCount,
  dayOptions,
  defaultScheduleConfig,
  defaultExcelPath,
  emptyAdminState,
  isPlayableMatch,
  matchNeedsSchedule,
  manualOverrideToAssignment,
  matchScheduleKey,
  matchTeamsLabel,
  normalizeScheduleConfig,
  parseCategoriesFromRows,
  publicStorageKey,
  scheduleToManualOverride,
  shufflePairsKeepingLocks,
  slugFromName,
  timeOptionsForDay,
  type AdminState,
  type CategoryData,
  type CategoryDrawSet,
  type DayKey,
  type Draw,
  type ManualPairLockMap,
  type ManualScheduleMap,
  type ManualScheduleOverride,
  type Match,
  type PairLockMap,
  type Pair,
  type PublishedTournament,
  type RestrictionRule,
  type RuleMode,
  type ScheduleAssignment,
  type ScheduleSummaryRow,
  type ScheduleWindow,
  type SelectionByCategory,
  type SheetCell,
  type Team,
  type TournamentScheduleConfig,
} from "../lib/tournament";

const adminPassword = "landerlander";
const tabs = [
  { id: "import", label: "Datos" },
  { id: "rules", label: "Restricciones" },
  { id: "draws", label: "Cuadros y horarios" },
  { id: "public", label: "Publicacion" },
] as const;

type AdminTab = (typeof tabs)[number]["id"];

function safeJsonParse<T>(value: string | null) {
  if (!value) return null;

  try {
    return JSON.parse(value) as T;
  } catch {
    return null;
  }
}

function readStoredAdminState(): AdminState {
  if (typeof window === "undefined") return emptyAdminState;

  return (
    safeJsonParse<AdminState>(window.localStorage.getItem(adminStorageKey)) ??
    emptyAdminState
  );
}

function workbookToCategories(workbook: XLSX.WorkBook) {
  const sheetName =
    workbook.SheetNames.find((name) =>
      name.toLocaleLowerCase("es-ES").includes("categoria"),
    ) ?? workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json<SheetCell[]>(sheet, {
    defval: "",
    header: 1,
  });

  return parseCategoriesFromRows(rows);
}

function subscribeClientReady(callback: () => void) {
  const timeout = window.setTimeout(callback, 0);

  return () => window.clearTimeout(timeout);
}

function getClientReadySnapshot() {
  return true;
}

function getServerReadySnapshot() {
  return false;
}

function useClientReady() {
  return useSyncExternalStore(
    subscribeClientReady,
    getClientReadySnapshot,
    getServerReadySnapshot,
  );
}

function AdminPasswordGate({ children }: { children: React.ReactNode }) {
  const clientReady = useClientReady();
  const [password, setPassword] = useState("");
  const [manualUnlock, setManualUnlock] = useState(false);
  const [error, setError] = useState("");
  const storedUnlock =
    clientReady &&
    typeof window !== "undefined" &&
    window.sessionStorage.getItem(adminAuthStorageKey) === "ok";
  const unlocked = manualUnlock || storedUnlock;

  function submit(event: React.FormEvent) {
    event.preventDefault();

    if (password === adminPassword) {
      window.sessionStorage.setItem(adminAuthStorageKey, "ok");
      setManualUnlock(true);
      setError("");
      return;
    }

    setError("Password incorrecta");
  }

  if (unlocked) return <>{children}</>;

  return (
    <main className="wc-app flex min-h-screen items-center justify-center p-6">
      <form className="wc-card wc-accent-top w-full max-w-sm p-6" onSubmit={submit}>
        <div className="mb-5 flex items-center gap-3">
          <span className="rg-logo-mark">
            <Trophy className="h-6 w-6" />
          </span>
          <div>
            <p className="wc-eyebrow">Padel Bracket</p>
            <h1 className="wc-title text-2xl">Panel privado</h1>
          </div>
        </div>
        <label className="wc-row-label mb-1.5 block">
          Acceso privado
        </label>
        <input
          className="wc-field h-11 w-full"
          onChange={(event) => setPassword(event.target.value)}
          placeholder="Contrasena"
          type="password"
          value={password}
        />
        {error ? (
          <p className="mt-2 text-sm font-bold text-[var(--coral-600)]">{error}</p>
        ) : null}
        <button className="wc-btn wc-btn-dark mt-4 h-11 w-full" type="submit">
          <LockKeyhole className="h-4 w-4" />
          Entrar al panel
        </button>
      </form>
    </main>
  );
}

export default function TournamentAdmin() {
  return (
    <AdminPasswordGate>
      <TournamentAdminClient />
    </AdminPasswordGate>
  );
}

function TournamentAdminClient() {
  const [storedState] = useState<AdminState>(() => readStoredAdminState());
  const [activeTab, setActiveTab] = useState<AdminTab>("import");
  const [activeCategoryId, setActiveCategoryId] = useState(
    storedState.activeCategoryId,
  );
  const [categories, setCategories] = useState<CategoryData[]>(
    storedState.categories,
  );
  const [previewCategories, setPreviewCategories] = useState<CategoryData[]>(
    [],
  );
  const [mainSelectionsByCategory, setMainSelectionsByCategory] =
    useState<SelectionByCategory>(storedState.mainSelectionsByCategory);
  const [consolationSelectionsByCategory, setConsolationSelectionsByCategory] =
    useState<SelectionByCategory>(storedState.consolationSelectionsByCategory);
  const [manualPairLocks, setManualPairLocks] = useState<ManualPairLockMap>(
    storedState.manualPairLocks ?? {},
  );
  const [manualScheduleOverrides, setManualScheduleOverrides] =
    useState<ManualScheduleMap>(storedState.manualScheduleOverrides ?? {});
  const [scheduleConfig, setScheduleConfig] = useState<TournamentScheduleConfig>(
    () =>
      normalizeScheduleConfig(
        storedState.scheduleConfig ?? defaultScheduleConfig,
      ),
  );
  const [publishedSlug, setPublishedSlug] = useState(storedState.publishedSlug);
  const [loadingExcel, setLoadingExcel] = useState(false);
  const [importError, setImportError] = useState("");
  const [copied, setCopied] = useState(false);
  const [remoteStatus, setRemoteStatus] = useState<
    "idle" | "saving" | "ok" | "error"
  >("idle");

  useEffect(() => {
    if (typeof window === "undefined") return;

    const state: AdminState = {
      activeCategoryId,
      categories,
      consolationSelectionsByCategory,
      mainSelectionsByCategory,
      manualPairLocks,
      manualScheduleOverrides,
      publishedSlug,
      scheduleConfig,
    };

    window.localStorage.setItem(adminStorageKey, JSON.stringify(state));
  }, [
    activeCategoryId,
    categories,
    consolationSelectionsByCategory,
    mainSelectionsByCategory,
    manualPairLocks,
    manualScheduleOverrides,
    publishedSlug,
    scheduleConfig,
  ]);

  const activeCategory =
    categories.find((category) => category.id === activeCategoryId) ??
    categories[0] ??
    null;
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
    drawSets.find((drawSet) => drawSet.categoryId === activeCategory?.id) ??
    drawSets[0] ??
    null;
  const schedule = useMemo(
    () => buildGlobalSchedule(drawSets, manualScheduleOverrides, scheduleConfig),
    [drawSets, manualScheduleOverrides, scheduleConfig],
  );
  const summaryRows = useMemo(
    () => buildScheduleSummaryRows(categories, drawSets, schedule.assignments),
    [categories, drawSets, schedule.assignments],
  );
  const reviewCount = countReviewPairs(categories);
  const manualPairLockCount = Object.values(manualPairLocks).reduce(
    (total, locks) => total + Object.keys(locks).length,
    0,
  );
  const manualScheduleCount = Object.keys(manualScheduleOverrides).length;
  const totalLockCount = manualPairLockCount + manualScheduleCount;
  const publicPath = publishedSlug ? `/publico/${publishedSlug}` : "";
  const publicUrl =
    typeof window !== "undefined" && publicPath
      ? `${window.location.origin}${publicPath}`
      : publicPath;

  async function loadPublicExcel() {
    setLoadingExcel(true);
    setImportError("");

    try {
      const response = await fetch(defaultExcelPath);

      if (!response.ok) throw new Error("No se pudo cargar el Excel de public");

      const buffer = await response.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const parsed = workbookToCategories(workbook);

      if (!parsed.length)
        throw new Error("No he detectado parejas en el Excel");

      setPreviewCategories(parsed);
      setActiveTab("import");
    } catch (error) {
      setImportError(
        error instanceof Error ? error.message : "Error cargando Excel",
      );
    } finally {
      setLoadingExcel(false);
    }
  }

  async function handleUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];

    if (!file) return;

    setLoadingExcel(true);
    setImportError("");

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const parsed = workbookToCategories(workbook);

      if (!parsed.length)
        throw new Error("No he detectado parejas en el Excel");

      setPreviewCategories(parsed);
      setActiveTab("import");
    } catch (error) {
      setImportError(
        error instanceof Error ? error.message : "Error cargando Excel",
      );
    } finally {
      setLoadingExcel(false);
      event.target.value = "";
    }
  }

  function applyPreview() {
    setCategories(previewCategories);
    setActiveCategoryId(previewCategories[0]?.id ?? "");
    setMainSelectionsByCategory({});
    setConsolationSelectionsByCategory({});
    setManualPairLocks({});
    setManualScheduleOverrides({});
    setPublishedSlug("");
    setPreviewCategories([]);
    setActiveTab("rules");
  }

  function updatePair(
    categoryId: string,
    pairId: string,
    patch: Partial<Pair>,
  ) {
    setCategories((current) =>
      current.map((category) =>
        category.id === categoryId
          ? {
              ...category,
              pairs: category.pairs.map((pair) =>
                pair.id === pairId ? { ...pair, ...patch } : pair,
              ),
            }
          : category,
      ),
    );
  }

  function updateRule(
    categoryId: string,
    pairId: string,
    ruleId: string,
    patch: Partial<RestrictionRule>,
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
                      rules: pair.rules.map((rule) =>
                        rule.id === ruleId ? { ...rule, ...patch } : rule,
                      ),
                    }
                  : pair,
              ),
            }
          : category,
      ),
    );
  }

  function addRule(categoryId: string, pairId: string) {
    updatePair(categoryId, pairId, {
      rules: [
        ...(categories
          .find((category) => category.id === categoryId)
          ?.pairs.find((pair) => pair.id === pairId)?.rules ?? []),
        {
          confidence: 100,
          day: "sabado",
          from: "08:30",
          id: `manual-rule-${Date.now().toString(36)}`,
          mode: "available",
          source: "manual",
          to: "21:30",
        },
      ],
    });
  }

  function removeRule(categoryId: string, pairId: string, ruleId: string) {
    const pair = categories
      .find((category) => category.id === categoryId)
      ?.pairs.find((item) => item.id === pairId);

    if (!pair) return;

    updatePair(categoryId, pairId, {
      rules: pair.rules.filter((rule) => rule.id !== ruleId),
    });
  }

  function movePairInBracket(
    categoryId: string,
    slotIndex: number,
    pairId: string,
  ) {
    const categoryLocks = manualPairLocks[categoryId] ?? {};
    const lockedSlots = new Set(
      Object.keys(categoryLocks).map((slot) => Number.parseInt(slot, 10)),
    );
    const lockedPairs = new Set(Object.values(categoryLocks));

    if (lockedSlots.has(slotIndex) || lockedPairs.has(pairId)) return;

    setCategories((current) =>
      current.map((category) => {
        if (category.id !== categoryId) return category;

        const selectedIndex = category.pairs.findIndex(
          (pair) => pair.id === pairId,
        );

        if (
          selectedIndex < 0 ||
          selectedIndex === slotIndex ||
          lockedSlots.has(selectedIndex)
        ) {
          return category;
        }

        const pairs = [...category.pairs];
        const selected = pairs[selectedIndex];

        pairs[selectedIndex] = pairs[slotIndex];
        pairs[slotIndex] = selected;

        return {
          ...category,
          pairs: pairs.map((pair, index) => ({ ...pair, seed: index + 1 })),
        };
      }),
    );
    clearCategoryResults(categoryId);
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

  function clearCategorySchedule(categoryId: string) {
    setManualScheduleOverrides((current) =>
      clearCategoryManualSchedules(current, categoryId),
    );
  }

  function toggleMatchupLock(categoryId: string, slotIndexes: number[]) {
    const category = categories.find((item) => item.id === categoryId);

    if (!category || !slotIndexes.length) return;

    setManualPairLocks((current) => {
      const currentLocks = current[categoryId] ?? {};
      const allLocked = slotIndexes.every((slotIndex) =>
        Boolean(currentLocks[String(slotIndex)]),
      );
      const nextCategoryLocks = { ...currentLocks };

      if (allLocked) {
        slotIndexes.forEach((slotIndex) => {
          delete nextCategoryLocks[String(slotIndex)];
        });
      } else {
        slotIndexes.forEach((slotIndex) => {
          const pair = category.pairs[slotIndex];

          if (pair) nextCategoryLocks[String(slotIndex)] = pair.id;
        });
      }

      const next = { ...current };

      if (Object.keys(nextCategoryLocks).length) {
        next[categoryId] = nextCategoryLocks;
      } else {
        delete next[categoryId];
      }

      return next;
    });
  }

  function recalculateKeepingManual() {
    setManualScheduleOverrides((current) => ({ ...current }));
  }

  function refreshCategoryDraws(categoryId: string) {
    clearCategoryResults(categoryId);
    setActiveCategoryId(categoryId);
    setActiveTab("draws");
  }

  function randomizeCategory(categoryId: string) {
    const lockedPairs = manualPairLocks[categoryId] ?? {};

    setCategories((current) => {
      const target = current.find((category) => category.id === categoryId);

      if (!target) return current;

      let bestCategories = current;
      let bestScore = Number.POSITIVE_INFINITY;

      for (let attempt = 0; attempt < 120; attempt += 1) {
        const shuffled = shufflePairsKeepingLocks(target.pairs, lockedPairs);
        const candidate = current.map((category) =>
          category.id === categoryId
            ? { ...category, pairs: shuffled }
            : category,
        );
        const candidateDrawSets = buildDrawSets(
          candidate,
          { ...mainSelectionsByCategory, [categoryId]: {} },
          { ...consolationSelectionsByCategory, [categoryId]: {} },
        );
        const candidateSchedule = buildGlobalSchedule(
          candidateDrawSets,
          manualScheduleOverrides,
          scheduleConfig,
        );
        const score =
          candidateSchedule.conflicts * 1000 - candidateSchedule.saturdayCount;

        if (score < bestScore) {
          bestScore = score;
          bestCategories = candidate;
        }
      }

      return bestCategories;
    });
    clearCategoryResults(categoryId);
  }

  function selectWinner(categoryId: string, match: Match, team: Team) {
    if (!isPlayableMatch(match)) return;

    const update =
      match.draw === "main"
        ? setMainSelectionsByCategory
        : setConsolationSelectionsByCategory;

    update((current) => ({
      ...current,
      [categoryId]: {
        ...(current[categoryId] ?? {}),
        [match.id]: team.id,
      },
    }));
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

  async function publishTournament() {
    const slug = publishedSlug || slugFromName("torneo padel");
    const tournament: PublishedTournament = {
      categories,
      consolationSelectionsByCategory,
      mainSelectionsByCategory,
      manualScheduleOverrides,
      name: "Torneo de padel",
      publishedAt: new Date().toISOString(),
      scheduleConfig,
      slug,
    };
    const payload = JSON.stringify(tournament);

    // Immediate local copy (works in this browser even before the store is set).
    window.localStorage.setItem(publicStorageKey(slug), payload);
    setPublishedSlug(slug);

    // Persist to the shared store so the public link works on every device and
    // survives closing the browser.
    setRemoteStatus("saving");
    try {
      const response = await fetch(`/api/public/${slug}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-admin-password": adminPassword,
        },
        body: payload,
      });

      setRemoteStatus(response.ok ? "ok" : "error");
    } catch {
      setRemoteStatus("error");
    }
  }

  async function copyPublicLink() {
    if (!publicUrl) return;

    await navigator.clipboard.writeText(publicUrl);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1600);
  }

  return (
    <main className="wc-app min-h-screen text-[var(--ink)]">
      <header className="wc-topbar">
        <div className="wc-topbar-inner">
          <div className="flex min-w-0 items-center gap-3">
            <span className="rg-logo-mark" aria-hidden="true">
              <Trophy className="h-6 w-6" />
            </span>
            <div className="min-w-0">
              <p className="wc-kicker">Tournament desk</p>
              <h1 className="wc-brand-title truncate">Padel Bracket</h1>
            </div>
          </div>

          <div className="wc-topbar-actions">
            <button
              className="wc-btn wc-btn-primary"
              disabled={loadingExcel}
              onClick={loadPublicExcel}
              type="button"
            >
              <FileSpreadsheet className="h-4 w-4" />
              Excel demo
            </button>
            <label className="wc-btn wc-btn-ghost cursor-pointer">
              <Upload className="h-4 w-4" />
              Subir Excel
              <input
                accept=".xlsx,.xls,.csv"
                className="hidden"
                onChange={handleUpload}
                type="file"
              />
            </label>
          </div>
        </div>
      </header>

      <div className="wc-shell">
        {importError ? (
          <Notice
            icon={AlertTriangle}
            tone="warning"
            title="No se pudo importar"
            text={importError}
          />
        ) : null}

        {previewCategories.length ? (
          <PreviewPanel
            categories={previewCategories}
            onCancel={() => setPreviewCategories([])}
            onValidate={applyPreview}
          />
        ) : null}

        {categories.length ? (
          <>
            <TournamentCommandCenter
              activeCategoryName={activeCategory?.name ?? "Sin categoria"}
              categories={categories}
              lockCount={totalLockCount}
              manualPairLockCount={manualPairLockCount}
              manualScheduleCount={manualScheduleCount}
              onGoTo={setActiveTab}
              published={Boolean(publishedSlug)}
              reviewCount={reviewCount}
              scheduleConflicts={schedule.conflicts}
              scheduleTotal={schedule.total}
              saturdayCount={schedule.saturdayCount}
            />

            <div className="wc-workspace">
              <aside className="wc-rail-panel">
                <AdminTabs
                  activeTab={activeTab}
                  categoriesCount={categories.length}
                  lockCount={totalLockCount}
                  onChange={setActiveTab}
                  reviewCount={reviewCount}
                  scheduleTotal={schedule.total}
                />
                <OperationalChecklist
                  lockCount={totalLockCount}
                  published={Boolean(publishedSlug)}
                  reviewCount={reviewCount}
                  scheduleConflicts={schedule.conflicts}
                  scheduleTotal={schedule.total}
                />
              </aside>

              <section className="wc-stage">
                {activeTab === "rules" ? (
                  <ScheduleConfigPanel
                    onChange={setScheduleConfig}
                    scheduleConfig={scheduleConfig}
                  />
                ) : null}

                {activeTab !== "import" && activeTab !== "public" ? (
                  <CategoryTabs
                    activeCategoryId={activeCategory?.id ?? ""}
                    categories={categories}
                    onSelect={setActiveCategoryId}
                  />
                ) : null}

                {activeTab === "import" ? (
                  <ImportHelp
                    categories={categories}
                    loadingExcel={loadingExcel}
                    onLoadPublicExcel={loadPublicExcel}
                  />
                ) : null}

                {activeTab === "rules" && activeCategory ? (
                  <RestrictionsPanel
                    category={activeCategory}
                    onAddRule={addRule}
                    onRemoveRule={removeRule}
                    onRefreshDraws={refreshCategoryDraws}
                    onUpdatePair={updatePair}
                    onUpdateRule={updateRule}
                  />
                ) : null}

                {activeTab === "draws" && activeCategory && activeDrawSet ? (
                  <DrawsPanel
                    category={activeCategory}
                    drawSet={activeDrawSet}
                    manualPairLocks={manualPairLocks[activeCategory.id] ?? {}}
                    manualScheduleOverrides={manualScheduleOverrides}
                    onClearSchedule={clearCategorySchedule}
                    onMovePair={movePairInBracket}
                    onRandomize={randomizeCategory}
                    onRecalculateKeepingManual={recalculateKeepingManual}
                    onResetManualSchedule={resetManualSchedule}
                    onSelectWinner={selectWinner}
                    onToggleMatchupLock={toggleMatchupLock}
                    onUpdateManualSchedule={updateManualSchedule}
                    scheduleConfig={scheduleConfig}
                    schedule={schedule.assignments}
                  />
                ) : null}

                {activeTab === "public" ? (
                  <PublicAdminPanel
                    copied={copied}
                    onCopy={copyPublicLink}
                    onPublish={publishTournament}
                    publicPath={publicPath}
                    publicUrl={publicUrl}
                    published={Boolean(publishedSlug)}
                    remoteStatus={remoteStatus}
                    rows={summaryRows}
                  />
                ) : null}

                {schedule.conflicts ? (
                  <Notice
                    icon={AlertTriangle}
                    tone="warning"
                    title="Horarios a revisar"
                    text={`${schedule.conflicts} partido(s) no respetan al 100% restricciones o pisan pista/hora.`}
                  />
                ) : null}
              </section>
            </div>
          </>
        ) : previewCategories.length ? null : (
          <EmptyState
            loading={loadingExcel}
            onLoadPublicExcel={loadPublicExcel}
          />
        )}
      </div>
    </main>
  );
}

function PreviewPanel({
  categories,
  onCancel,
  onValidate,
}: {
  categories: CategoryData[];
  onCancel: () => void;
  onValidate: () => void;
}) {
  return (
    <section className="wc-card wc-accent-top p-4 sm:p-5">
      <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <p className="wc-eyebrow text-[var(--gold-600)]">Vista previa del Excel</p>
          <h2 className="wc-title mt-1 text-2xl">
            {categories.length} categorias / {countPairs(categories)} parejas
          </h2>
          <p className="mt-1 text-sm font-medium text-[var(--ink-soft)]">
            Revisa lo que se va a cargar. Las restricciones ya vienen
            autocompletadas en formato estructurado.
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          <button className="wc-btn wc-btn-ghost" onClick={onCancel} type="button">
            Cancelar
          </button>
          <button
            className="wc-btn wc-btn-primary"
            onClick={onValidate}
            type="button"
          >
            <CheckCircle2 className="h-4 w-4" />
            Validar carga
          </button>
        </div>
      </div>
      <div className="mt-4 overflow-x-auto">
        <table className="w-full min-w-[720px] border-collapse text-sm">
          <thead>
            <tr className="border-b border-[var(--line-strong)] text-left font-display text-xs uppercase text-[var(--ink-faint)]">
              <th className="px-3 py-2.5 font-semibold">Categoria</th>
              <th className="px-3 py-2.5 font-semibold">Parejas</th>
              <th className="px-3 py-2.5 font-semibold">Reglas auto</th>
              <th className="px-3 py-2.5 font-semibold">A revisar</th>
            </tr>
          </thead>
          <tbody>
            {categories.map((category) => {
              const review = category.pairs.filter(
                (pair) => pair.review,
              ).length;
              const rules = category.pairs.reduce(
                (total, pair) => total + pair.rules.length,
                0,
              );

              return (
                <tr
                  className="border-b border-[var(--line)] last:border-0"
                  key={category.id}
                >
                  <td className="px-3 py-2.5">
                    <span className="flex items-center gap-2 font-semibold">
                      <span
                        className="wc-crest"
                        style={{ backgroundColor: categoryCrestColor(category.id) }}
                      />
                      {category.name}
                    </span>
                  </td>
                  <td className="px-3 py-2.5 tabular font-semibold">
                    {category.pairs.length}
                  </td>
                  <td className="px-3 py-2.5 tabular text-[var(--ink-soft)]">
                    {rules}
                  </td>
                  <td className="px-3 py-2.5">
                    <span
                      className={`wc-chip ${review ? "wc-chip-amber" : "wc-chip-green"}`}
                    >
                      {review || "OK"}
                    </span>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function EmptyState({
  loading,
  onLoadPublicExcel,
}: {
  loading: boolean;
  onLoadPublicExcel: () => void;
}) {
  return (
    <section className="wc-onboarding">
      <div className="wc-onboarding-main">
        <p className="wc-kicker">Nuevo torneo</p>
        <h2 className="wc-onboarding-title">Del Excel al cuadro publico</h2>
        <p className="wc-onboarding-copy">
          Carga categorias y parejas, convierte las notas horarias en reglas
          editables, calcula cruces con prioridad por dia y deja bloqueado lo
          que el director del torneo ya haya decidido.
        </p>
        <button
          className="wc-btn wc-btn-dark mt-5 h-11"
          disabled={loading}
          onClick={onLoadPublicExcel}
          type="button"
        >
          <FileSpreadsheet className="h-4 w-4" />
          {loading ? "Cargando..." : "Cargar Excel de ejemplo"}
        </button>
      </div>

      <div className="wc-onboarding-flow">
        {[
          {
            label: "1",
            title: "Importar",
            text: "Detecta categorias, parejas y notas del Excel.",
          },
          {
            label: "2",
            title: "Validar",
            text: "Revisa SOLO PUEDE / NO PUEDE antes de sortear.",
          },
          {
            label: "3",
            title: "Fijar",
            text: "Bloquea cruces u horarios y recalcula alrededor.",
          },
          {
            label: "4",
            title: "Publicar",
            text: "Entrega horarios y cuadros en una vista para jugadores.",
          },
        ].map((item) => (
          <article className="wc-onboarding-step" key={item.label}>
            <span>{item.label}</span>
            <div>
              <h3>{item.title}</h3>
              <p>{item.text}</p>
            </div>
          </article>
        ))}
      </div>
    </section>
  );
}

function ImportHelp({
  categories,
  loadingExcel,
  onLoadPublicExcel,
}: {
  categories: CategoryData[];
  loadingExcel: boolean;
  onLoadPublicExcel: () => void;
}) {
  return (
    <section className="wc-card wc-accent-top p-4 sm:p-5">
      <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <p className="wc-eyebrow">Datos cargados</p>
          <h2 className="wc-title mt-1 text-2xl">Plantilla del torneo</h2>
          <p className="mt-1 text-sm font-medium text-[var(--ink-soft)]">
            El Excel activo tiene {categories.length} categorias y{" "}
            {countPairs(categories)} parejas. Puedes recargar el de ejemplo o
            subir uno nuevo desde la cabecera.
          </p>
        </div>
        <button
          className="wc-btn wc-btn-ghost"
          disabled={loadingExcel}
          onClick={onLoadPublicExcel}
          type="button"
        >
          <RotateCcw className="h-4 w-4" />
          Recargar Excel
        </button>
      </div>
      <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
        {categories.map((category) => {
          const review = category.pairs.filter((pair) => pair.review).length;

          return (
            <article className="wc-inset p-3.5" key={category.id}>
              <div className="flex items-center gap-2">
                <span
                  className="wc-crest"
                  style={{ backgroundColor: categoryCrestColor(category.id) }}
                />
                <h3 className="font-display text-base font-semibold uppercase tracking-wide">
                  {category.name}
                </h3>
              </div>
              <div className="mt-2.5 flex flex-wrap gap-1.5">
                <span className="wc-chip wc-chip-neutral">
                  <Users className="h-3.5 w-3.5" />
                  {category.pairs.length} parejas
                </span>
                <span
                  className={`wc-chip ${review ? "wc-chip-amber" : "wc-chip-green"}`}
                >
                  {review ? `${review} a revisar` : "Sin avisos"}
                </span>
              </div>
            </article>
          );
        })}
      </div>
    </section>
  );
}

function TournamentCommandCenter({
  activeCategoryName,
  categories,
  lockCount,
  manualPairLockCount,
  manualScheduleCount,
  onGoTo,
  published,
  reviewCount,
  saturdayCount,
  scheduleConflicts,
  scheduleTotal,
}: {
  activeCategoryName: string;
  categories: CategoryData[];
  lockCount: number;
  manualPairLockCount: number;
  manualScheduleCount: number;
  onGoTo: (tab: AdminTab) => void;
  published: boolean;
  reviewCount: number;
  saturdayCount: number;
  scheduleConflicts: number;
  scheduleTotal: number;
}) {
  const pairCount = countPairs(categories);
  const nextStep =
    reviewCount > 0
      ? {
          action: "Abrir restricciones",
          copy: `${reviewCount} pareja(s) necesitan una lectura humana antes de sortear.`,
          tab: "rules" as const,
          title: "Validar restricciones",
        }
      : scheduleConflicts > 0
        ? {
            action: "Resolver horarios",
            copy: `${scheduleConflicts} partido(s) chocan con reglas, pista u hora.`,
            tab: "draws" as const,
            title: "Limpiar conflictos",
          }
        : scheduleTotal === 0
          ? {
              action: "Ir a cuadros",
              copy: "El Excel esta cargado. Falta generar, fijar y revisar cruces.",
              tab: "draws" as const,
              title: "Construir el cuadro",
            }
          : !published
            ? {
                action: "Preparar publicacion",
                copy: "El calendario ya tiene partidos. Revisa candados y publica.",
                tab: "public" as const,
                title: "Listo para publicar",
              }
            : {
                action: "Actualizar publico",
                copy: "La salida publica existe. Cualquier cambio manual puede republicarse.",
                tab: "public" as const,
                title: "Torneo en vivo",
              };

  return (
    <section className="wc-command-center">
      <div className="wc-command-main">
        <div>
          <p className="wc-kicker">Estado operativo</p>
          <h2 className="wc-command-title">{nextStep.title}</h2>
          <p className="wc-command-copy">{nextStep.copy}</p>
        </div>
        <button
          className="wc-btn wc-btn-dark"
          onClick={() => onGoTo(nextStep.tab)}
          type="button"
        >
          {nextStep.action}
        </button>
      </div>

      <div className="wc-command-grid">
        <CommandMetric
          label="Categorias"
          meta={activeCategoryName}
          value={categories.length}
        />
        <CommandMetric label="Parejas" meta="Excel validado" value={pairCount} />
        <CommandMetric
          label="Partidos"
          meta={`${saturdayCount} en sabado`}
          value={scheduleTotal}
        />
        <CommandMetric
          label="Candados"
          meta={`${manualPairLockCount} cruces / ${manualScheduleCount} horarios`}
          value={lockCount}
        />
      </div>
    </section>
  );
}

function CommandMetric({
  label,
  meta,
  value,
}: {
  label: string;
  meta: string;
  value: number | string;
}) {
  return (
    <article className="wc-command-metric">
      <p className="wc-command-metric-label">{label}</p>
      <p className="wc-command-metric-value">{value}</p>
      <p className="wc-command-metric-meta">{meta}</p>
    </article>
  );
}

function OperationalChecklist({
  lockCount,
  published,
  reviewCount,
  scheduleConflicts,
  scheduleTotal,
}: {
  lockCount: number;
  published: boolean;
  reviewCount: number;
  scheduleConflicts: number;
  scheduleTotal: number;
}) {
  const checks = [
    {
      done: reviewCount === 0,
      label: reviewCount ? `${reviewCount} restricciones pendientes` : "Restricciones claras",
    },
    {
      done: scheduleTotal > 0,
      label: scheduleTotal ? `${scheduleTotal} partidos calculados` : "Cuadros por generar",
    },
    {
      done: scheduleConflicts === 0 && scheduleTotal > 0,
      label: scheduleConflicts ? `${scheduleConflicts} conflictos` : "Horarios coherentes",
    },
    {
      done: lockCount > 0,
      label: lockCount ? `${lockCount} decisiones fijadas` : "Sin candados manuales",
    },
    {
      done: published,
      label: published ? "Publicado" : "Pendiente de publicar",
    },
  ];

  return (
    <section className="wc-readiness">
      <p className="wc-kicker">Checklist</p>
      <div className="mt-3 grid gap-2">
        {checks.map((check) => (
          <div className="wc-readiness-row" key={check.label}>
            <span
              className={`wc-readiness-dot ${
                check.done ? "wc-readiness-dot-ok" : ""
              }`}
            />
            <span>{check.label}</span>
          </div>
        ))}
      </div>
    </section>
  );
}

function AdminTabs({
  activeTab,
  categoriesCount,
  lockCount,
  onChange,
  reviewCount,
  scheduleTotal,
}: {
  activeTab: AdminTab;
  categoriesCount: number;
  lockCount: number;
  onChange: (tab: AdminTab) => void;
  reviewCount: number;
  scheduleTotal: number;
}) {
  const tabDetails: Record<
    AdminTab,
    { index: string; status: string; done: boolean }
  > = {
    import: {
      done: categoriesCount > 0,
      index: "1",
      status: categoriesCount ? `${categoriesCount} categorias` : "Pendiente",
    },
    rules: {
      done: categoriesCount > 0 && reviewCount === 0,
      index: "2",
      status: reviewCount ? `${reviewCount} a revisar` : "Validado",
    },
    draws: {
      done: scheduleTotal > 0,
      index: "3",
      status: scheduleTotal ? `${scheduleTotal} partidos` : "Pendiente",
    },
    public: {
      done: scheduleTotal > 0,
      index: "4",
      status: lockCount ? `${lockCount} candados` : "Publicacion",
    },
  };

  return (
    <nav className="wc-process" aria-label="Fases del torneo">
      {tabs.map((tab) => {
        const active = activeTab === tab.id;
        const done = tabDetails[tab.id].done;

        return (
          <button
            aria-current={active ? "step" : undefined}
            className={`wc-process-step ${active ? "wc-process-step-active" : ""}`}
            key={tab.id}
            onClick={() => onChange(tab.id)}
            type="button"
          >
            <span
              className={`wc-process-badge ${
                active
                  ? "wc-process-badge-active"
                  : done
                    ? "wc-process-badge-done"
                    : "wc-process-badge-idle"
              }`}
            >
              {done && !active ? (
                <CheckCircle2 className="h-5 w-5" />
              ) : (
                tabDetails[tab.id].index
              )}
            </span>
            <span className="min-w-0">
              <span
                className={`wc-process-name ${
                  active ? "text-[var(--pitch-900)]" : "text-[var(--ink)]"
                }`}
              >
                {tab.label}
              </span>
              <span
                className={`wc-process-status block truncate ${
                  active ? "text-[var(--clay-dark)]" : "text-[var(--ink-soft)]"
                }`}
              >
                {tabDetails[tab.id].status}
              </span>
            </span>
          </button>
        );
      })}
    </nav>
  );
}

const categoryCrestPalette = [
  "#0d8a52",
  "#2563eb",
  "#d99e0b",
  "#e11d48",
  "#7c3aed",
  "#0891b2",
  "#ea580c",
  "#0f766e",
];

function categoryCrestColor(categoryId: string) {
  const total = Array.from(categoryId).reduce(
    (sum, character) => sum + character.charCodeAt(0),
    0,
  );

  return categoryCrestPalette[total % categoryCrestPalette.length];
}

function CategoryTabs({
  activeCategoryId,
  categories,
  onSelect,
}: {
  activeCategoryId: string;
  categories: CategoryData[];
  onSelect: (categoryId: string) => void;
}) {
  return (
    <section className="wc-card-flat p-2.5">
      <div className="flex flex-wrap gap-2">
        {categories.map((category) => {
          const active = category.id === activeCategoryId;

          return (
            <button
              className={`wc-pill ${active ? "wc-pill-active" : ""}`}
              key={category.id}
              onClick={() => onSelect(category.id)}
              type="button"
            >
              <span
                className="wc-crest"
                style={{ backgroundColor: categoryCrestColor(category.id) }}
              />
              {category.name}
              <span className="wc-pill-count">{category.pairs.length}</span>
            </button>
          );
        })}
      </div>
    </section>
  );
}

function RestrictionsPanel({
  category,
  onAddRule,
  onRefreshDraws,
  onRemoveRule,
  onUpdatePair,
  onUpdateRule,
}: {
  category: CategoryData;
  onAddRule: (categoryId: string, pairId: string) => void;
  onRefreshDraws: (categoryId: string) => void;
  onRemoveRule: (categoryId: string, pairId: string, ruleId: string) => void;
  onUpdatePair: (
    categoryId: string,
    pairId: string,
    patch: Partial<Pair>,
  ) => void;
  onUpdateRule: (
    categoryId: string,
    pairId: string,
    ruleId: string,
    patch: Partial<RestrictionRule>,
  ) => void;
}) {
  const reviewPairs = category.pairs.filter((pair) => pair.review).length;
  const structuredRules = category.pairs.reduce(
    (total, pair) => total + pair.rules.length,
    0,
  );

  return (
    <section className="wc-pane">
      <div className="wc-pane-head">
        <div>
          <p className="wc-kicker">Restricciones</p>
          <h2 className="wc-pane-title">{category.name}</h2>
          <p className="wc-pane-copy">
            Traduce las notas del Excel a reglas horarias reales. Estas reglas
            se combinan con el motor horario global del torneo.
          </p>
        </div>
        <div className="wc-pane-actions">
          <span className="wc-chip wc-chip-neutral">
            {structuredRules} reglas
          </span>
          <span
            className={`wc-chip ${reviewPairs ? "wc-chip-amber" : "wc-chip-green"}`}
          >
            {reviewPairs ? `${reviewPairs} a revisar` : "Sin avisos"}
          </span>
        </div>
        <button
          className="wc-btn wc-btn-primary"
          onClick={() => onRefreshDraws(category.id)}
          type="button"
        >
          <RotateCcw className="h-4 w-4" />
          Validar y recalcular
        </button>
      </div>

      <div className="wc-rules-layout">
        <div className="wc-rules-list">
          {category.pairs.map((pair) => (
            <article
              className={`wc-pair-row ${pair.review ? "wc-pair-row-review" : ""}`}
              key={pair.id}
            >
              <div className="wc-pair-identity">
                <span className="wc-seed">{pair.seed}</span>
                <div className="min-w-0">
                  <p className="wc-row-label">Pareja</p>
                  <h3>{pair.playerOne} / {pair.playerTwo}</h3>
                </div>
                <label className="wc-review-toggle">
                  <input
                    checked={pair.review}
                    className="h-4 w-4 accent-[var(--amber-500)]"
                    onChange={(event) =>
                      onUpdatePair(category.id, pair.id, {
                        review: event.target.checked,
                      })
                    }
                    type="checkbox"
                  />
                  Revisar
                </label>
              </div>

              <div className="wc-excel-note">
                <p className="wc-row-label">Excel</p>
                <p>{pair.rawRestriction || "Sin restriccion compacta"}</p>
                {pair.rawNotes ? <small>{pair.rawNotes}</small> : null}
                {pair.reviewReasons.length ? (
                  <div className="mt-2 flex flex-wrap gap-1.5">
                    {pair.reviewReasons.map((reason) => (
                      <span className="wc-chip wc-chip-amber" key={reason}>
                        {reason}
                      </span>
                    ))}
                  </div>
                ) : null}
              </div>

              <div className="wc-rule-stack">
                <div className="wc-rule-stack-head">
                  <p className="wc-row-label">Reglas para calcular</p>
                  <button
                    className="wc-btn wc-btn-ghost wc-btn-sm"
                    onClick={() => onAddRule(category.id, pair.id)}
                    type="button"
                  >
                    <Plus className="h-3.5 w-3.5" />
                    Regla
                  </button>
                </div>

                {pair.rules.length ? (
                  pair.rules.map((rule) => (
                    <RuleEditor
                      key={rule.id}
                      onRemove={() =>
                        onRemoveRule(category.id, pair.id, rule.id)
                      }
                      onUpdate={(patch) =>
                        onUpdateRule(category.id, pair.id, rule.id, patch)
                      }
                      rule={rule}
                    />
                  ))
                ) : (
                  <div className="wc-empty-inline">
                    Sin reglas: disponible en cualquier hueco del torneo.
                  </div>
                )}
              </div>
            </article>
          ))}
        </div>
      </div>
    </section>
  );
}

function ScheduleConfigPanel({
  onChange,
  scheduleConfig,
}: {
  onChange: (config: TournamentScheduleConfig) => void;
  scheduleConfig: TournamentScheduleConfig;
}) {
  const config = normalizeScheduleConfig(scheduleConfig);
  const priorityLabels = ["Mayoria", "Luego", "Ultimo recurso"];

  function updateWindow(
    day: DayKey,
    index: number,
    patch: Partial<ScheduleWindow>,
  ) {
    const windows = [...config.dayWindows[day]];

    windows[index] = {
      ...windows[index],
      ...patch,
    };

    onChange({
      ...config,
      dayWindows: {
        ...config.dayWindows,
        [day]: windows,
      },
    });
  }

  function updatePriority(index: number, day: DayKey) {
    const nextPriority = [...config.dayPriority];
    const currentIndex = nextPriority.indexOf(day);

    if (currentIndex >= 0) {
      const currentDay = nextPriority[index];

      nextPriority[index] = day;
      nextPriority[currentIndex] = currentDay;
    } else {
      nextPriority[index] = day;
    }

    onChange({
      ...config,
      dayPriority: nextPriority,
    });
  }

  return (
    <section className="wc-schedule-config">
      <div className="flex flex-col gap-3 xl:flex-row xl:items-start xl:justify-between">
        <div className="min-w-0">
          <p className="wc-kicker">Motor horario global</p>
          <h3 className="wc-section-title">Dias y ventanas del torneo</h3>
          <p className="mt-1 text-sm font-medium text-[var(--ink-soft)]">
            Estos tramos se aplican a todas las categorias al validar, sortear
            y recalcular con las 8 pistas compartidas.
          </p>
        </div>
        <button
          className="wc-btn wc-btn-ghost wc-btn-sm"
          onClick={() => onChange(defaultScheduleConfig)}
          type="button"
        >
          <RotateCcw className="h-4 w-4" />
          Restaurar defecto
        </button>
      </div>

      <div className="mt-3 grid gap-3">
        <div className="grid gap-2 lg:grid-cols-3">
          {dayOptions.map((day) => (
            <article
              className="wc-day-window-card"
              key={day.key}
            >
              <div className="mb-2 flex items-center justify-between gap-2">
                <h4 className="wc-row-label">
                  {day.label}
                </h4>
                <span className="wc-chip wc-chip-neutral">
                  {config.dayWindows[day.key].length} tramo
                  {config.dayWindows[day.key].length === 1 ? "" : "s"}
                </span>
              </div>
              <div className="grid gap-2">
                {config.dayWindows[day.key].map((window, index) => (
                  <div
                    className="grid grid-cols-[auto_1fr_1fr] items-center gap-2"
                    key={`${day.key}-${index}`}
                  >
                    <span className="font-display text-[0.7rem] font-bold uppercase text-[var(--ink-faint)]">
                      T{index + 1}
                    </span>
                    <input
                      className="wc-field h-9 min-w-0 tabular"
                      onChange={(event) =>
                        updateWindow(day.key, index, {
                          from: event.target.value,
                        })
                      }
                      type="time"
                      value={window.from}
                    />
                    <input
                      className="wc-field h-9 min-w-0 tabular"
                      onChange={(event) =>
                        updateWindow(day.key, index, {
                          to: event.target.value,
                        })
                      }
                      type="time"
                      value={window.to}
                    />
                  </div>
                ))}
              </div>
            </article>
          ))}
        </div>

        <article className="wc-day-window-card">
          <p className="wc-row-label">
            Prioridad automatica
          </p>
          <div className="mt-2 grid gap-2">
            {config.dayPriority.map((day, index) => (
              <label className="grid gap-1" key={`${day}-${index}`}>
                <span className="text-xs font-bold text-[var(--ink-soft)]">
                  {priorityLabels[index] ?? `Prioridad ${index + 1}`}
                </span>
                <select
                  className="wc-field h-9"
                  onChange={(event) =>
                    updatePriority(index, event.target.value as DayKey)
                  }
                  value={day}
                >
                  {dayOptions.map((option) => (
                    <option key={option.key} value={option.key}>
                      {option.label}
                    </option>
                  ))}
                </select>
              </label>
            ))}
          </div>
        </article>
      </div>
    </section>
  );
}

function RuleEditor({
  onRemove,
  onUpdate,
  rule,
}: {
  onRemove: () => void;
  onUpdate: (patch: Partial<RestrictionRule>) => void;
  rule: RestrictionRule;
}) {
  const available = rule.mode === "available";

  return (
    <div
      className={`wc-rule-editor ${
        available ? "wc-rule-editor-available" : "wc-rule-editor-blocked"
      }`}
    >
      <select
        className={`wc-field h-9 ${
          available
            ? "bg-[var(--pitch-50)] text-[var(--pitch-800)]"
            : "bg-[var(--coral-50)] text-[var(--coral-600)]"
        }`}
        onChange={(event) => onUpdate({ mode: event.target.value as RuleMode })}
        value={rule.mode}
      >
        <option value="available">SOLO PUEDE</option>
        <option value="blocked">NO PUEDE</option>
      </select>
      <select
        className="wc-field h-9"
        onChange={(event) => onUpdate({ day: event.target.value as DayKey })}
        value={rule.day}
      >
        {dayOptions.map((day) => (
          <option key={day.key} value={day.key}>
            {day.label}
          </option>
        ))}
      </select>
      <input
        className="wc-field h-9 tabular"
        onChange={(event) => onUpdate({ from: event.target.value })}
        type="time"
        value={rule.from}
      />
      <input
        className="wc-field h-9 tabular"
        onChange={(event) => onUpdate({ to: event.target.value })}
        type="time"
        value={rule.to}
      />
      <button
        aria-label="Eliminar regla"
        className="wc-icon-btn h-9 w-9 hover:border-[var(--coral-500)] hover:bg-[var(--coral-50)] hover:text-[var(--coral-600)]"
        onClick={onRemove}
        type="button"
      >
        <Trash2 className="h-4 w-4" />
      </button>
    </div>
  );
}

function DrawsPanel({
  category,
  drawSet,
  manualPairLocks,
  manualScheduleOverrides,
  onClearSchedule,
  onMovePair,
  onRandomize,
  onRecalculateKeepingManual,
  onResetManualSchedule,
  onSelectWinner,
  onToggleMatchupLock,
  onUpdateManualSchedule,
  scheduleConfig,
  schedule,
}: {
  category: CategoryData;
  drawSet: CategoryDrawSet;
  manualPairLocks: PairLockMap;
  manualScheduleOverrides: ManualScheduleMap;
  onClearSchedule: (categoryId: string) => void;
  onMovePair: (categoryId: string, slotIndex: number, pairId: string) => void;
  onRandomize: (categoryId: string) => void;
  onRecalculateKeepingManual: () => void;
  onResetManualSchedule: (scheduleKey: string) => void;
  onSelectWinner: (categoryId: string, match: Match, team: Team) => void;
  onToggleMatchupLock: (categoryId: string, slotIndexes: number[]) => void;
  onUpdateManualSchedule: (
    scheduleKey: string,
    override: ManualScheduleOverride,
  ) => void;
  scheduleConfig: TournamentScheduleConfig;
  schedule: Record<string, ScheduleAssignment>;
}) {
  const lockedCruces = Object.keys(manualPairLocks).length;
  const lockedHorarios = Object.keys(manualScheduleOverrides).filter((key) =>
    key.startsWith(`${category.id}:`),
  ).length;

  return (
    <div className="grid gap-5">
      <section className="wc-pane">
        <div className="wc-pane-head">
          <div className="flex items-center gap-2.5">
            <span
              className="wc-crest"
              style={{ backgroundColor: categoryCrestColor(category.id) }}
            />
            <div>
              <p className="wc-kicker">Mesa de cuadros</p>
              <h2 className="wc-pane-title">{category.name}</h2>
              <p className="wc-pane-copy">
                Sortea, fija cruces concretos y bloquea horarios manuales antes
                de recalcular el resto.
              </p>
            </div>
          </div>
          <div className="wc-pane-actions">
            <span className="wc-chip wc-chip-neutral">
              {lockedCruces} cruces fijos
            </span>
            <span className="wc-chip wc-chip-neutral">
              {lockedHorarios} horarios fijos
            </span>
          </div>
          <button
            className="wc-btn wc-btn-primary"
            onClick={() => onRandomize(category.id)}
            type="button"
          >
            <Shuffle className="h-4 w-4" />
            Sortear cuadro
          </button>
        </div>
      </section>

      <ScheduleEditor
        categoryId={category.id}
        drawSet={drawSet}
        manualScheduleOverrides={manualScheduleOverrides}
        onClearSchedule={onClearSchedule}
        onRecalculateKeepingManual={onRecalculateKeepingManual}
        onResetManualSchedule={onResetManualSchedule}
        onUpdateManualSchedule={onUpdateManualSchedule}
        scheduleConfig={scheduleConfig}
        schedule={schedule}
      />

      <MatchupEditor
        category={category}
        manualPairLocks={manualPairLocks}
        onMovePair={onMovePair}
        onToggleLock={onToggleMatchupLock}
      />

      <section className="wc-bracket-field grid gap-5">
        <BracketView
          categoryId={category.id}
          draw={drawSet.mainDraw}
          onSelectWinner={onSelectWinner}
          schedule={schedule}
          title="Cuadro principal"
        />
        <BracketView
          categoryId={category.id}
          draw={drawSet.consolationDraw}
          onSelectWinner={onSelectWinner}
          schedule={schedule}
          title="Consolacion"
        />
      </section>
    </div>
  );
}

function MatchupEditor({
  category,
  manualPairLocks,
  onMovePair,
  onToggleLock,
}: {
  category: CategoryData;
  manualPairLocks: PairLockMap;
  onMovePair: (categoryId: string, slotIndex: number, pairId: string) => void;
  onToggleLock: (categoryId: string, slotIndexes: number[]) => void;
}) {
  const lockedPairIds = new Set(Object.values(manualPairLocks));
  const [draggingPairId, setDraggingPairId] = useState<string | null>(null);
  const [dragOverSlot, setDragOverSlot] = useState<number | null>(null);
  const matchupSlots = Array.from(
    { length: Math.ceil(category.pairs.length / 2) },
    (_, index) => [index * 2, index * 2 + 1],
  );

  function canDropOnSlot(slot: number, pairId: string | null) {
    if (!pairId) return false;
    if (manualPairLocks[String(slot)]) return false;

    const currentPair = category.pairs[slot];

    return !lockedPairIds.has(pairId) || currentPair?.id === pairId;
  }

  return (
    <section className="wc-pane">
      <div className="wc-pane-head">
        <div>
          <p className="wc-kicker">Cruces manuales</p>
          <h3 className="wc-section-title">Primera ronda</h3>
          <p className="wc-pane-copy">
            Mueve parejas entre slots y bloquea los cruces que no deben cambiar
            en futuros sorteos.
          </p>
        </div>
        <span
          className={`wc-chip ${
            Object.keys(manualPairLocks).length
              ? "wc-chip-green"
              : "wc-chip-neutral"
          }`}
        >
          <LockKeyhole className="h-3.5 w-3.5" />
          {Object.keys(manualPairLocks).length
            ? `${Object.keys(manualPairLocks).length} slots protegidos`
            : "Sin cruces bloqueados"}
        </span>
      </div>
      <div className="wc-matchup-grid">
        {matchupSlots.map((slots, matchupIndex) => {
          const lockableSlots = slots.filter((slot) => category.pairs[slot]);
          const locked = lockableSlots.length > 0 && lockableSlots.every((slot) =>
            Boolean(manualPairLocks[String(slot)]),
          );

          return (
            <article
              className={`wc-matchup-card ${
                locked
                  ? "wc-matchup-card-locked"
                  : ""
              }`}
              key={slots.join("-")}
            >
              <div className="mb-3 flex items-center justify-between gap-2">
                <p className="wc-row-label">
                  Partido {matchupIndex + 1}
                </p>
                <button
                  aria-label={
                    locked
                      ? `Desbloquear partido ${matchupIndex + 1}`
                      : `Bloquear partido ${matchupIndex + 1}`
                  }
                  className={`inline-flex h-9 w-9 items-center justify-center rounded-lg border transition ${
                    locked
                      ? "border-[var(--pitch-600)] bg-[var(--pitch-600)] text-white hover:bg-[var(--pitch-700)]"
                      : "border-[var(--line-strong)] bg-[var(--surface)] text-[var(--ink-soft)] hover:border-[var(--pitch-500)] hover:text-[var(--pitch-700)]"
                  }`}
                  onClick={() => onToggleLock(category.id, lockableSlots)}
                  title={locked ? "Desbloquear cruce" : "Bloquear cruce"}
                  type="button"
                >
                  {locked ? (
                    <LockKeyhole className="h-4 w-4" />
                  ) : (
                    <LockOpen className="h-4 w-4" />
                  )}
                </button>
              </div>

              <div className="grid gap-2">
                {slots.map((slot) => {
                  const pair = category.pairs[slot];
                  const slotLocked = Boolean(manualPairLocks[String(slot)]);
                  const canDrop = canDropOnSlot(slot, draggingPairId);

                  return (
                    <div
                      className={`wc-matchup-slot ${
                        dragOverSlot === slot && canDrop
                          ? "wc-matchup-slot-over"
                          : ""
                      }`}
                      key={slot}
                      onDragLeave={() => {
                        if (dragOverSlot === slot) setDragOverSlot(null);
                      }}
                      onDragOver={(event) => {
                        if (!canDrop) return;

                        event.preventDefault();
                        event.dataTransfer.dropEffect = "move";
                        setDragOverSlot(slot);
                      }}
                      onDrop={(event) => {
                        event.preventDefault();
                        const pairId = event.dataTransfer.getData(
                          "application/x-padel-pair",
                        );

                        setDragOverSlot(null);
                        setDraggingPairId(null);

                        if (!canDropOnSlot(slot, pairId)) return;

                        onMovePair(category.id, slot, pairId);
                      }}
                    >
                      <span className="wc-row-label">
                        Slot {slot + 1}
                      </span>
                      {pair ? (
                        <>
                          <div
                            className={`wc-draggable-pair ${
                              slotLocked ? "wc-draggable-pair-locked" : ""
                            }`}
                            draggable={!slotLocked && !lockedPairIds.has(pair.id)}
                            onDragEnd={() => {
                              setDraggingPairId(null);
                              setDragOverSlot(null);
                            }}
                            onDragStart={(event) => {
                              if (slotLocked || lockedPairIds.has(pair.id)) {
                                event.preventDefault();
                                return;
                              }

                              event.dataTransfer.effectAllowed = "move";
                              event.dataTransfer.setData(
                                "application/x-padel-pair",
                                pair.id,
                              );
                              event.dataTransfer.setData("text/plain", pair.id);
                              setDraggingPairId(pair.id);
                            }}
                            title={
                              slotLocked
                                ? "Slot bloqueado"
                                : "Arrastra para mover esta pareja"
                            }
                          >
                            <span className="wc-drag-handle" aria-hidden="true">
                              ::
                            </span>
                            <span className="min-w-0 truncate">
                              {pair.seed}. {pair.playerOne} / {pair.playerTwo}
                            </span>
                          </div>
                          <select
                            aria-label={`Cambiar slot ${slot + 1}`}
                            className="wc-field h-9 w-full min-w-0 truncate"
                            disabled={slotLocked}
                            onChange={(event) =>
                              onMovePair(category.id, slot, event.target.value)
                            }
                            value={pair.id}
                          >
                            {category.pairs.map((option) => (
                              <option
                                disabled={
                                  lockedPairIds.has(option.id) &&
                                  option.id !== pair.id
                                }
                                key={option.id}
                                value={option.id}
                              >
                                {option.seed}. {option.playerOne} /{" "}
                                {option.playerTwo}
                              </option>
                            ))}
                          </select>
                        </>
                      ) : (
                        <div
                          className={`wc-bye-slot ${
                            dragOverSlot === slot && canDrop
                              ? "wc-matchup-slot-over"
                              : ""
                          }`}
                        >
                          Bye
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </article>
          );
        })}
      </div>
    </section>
  );
}

function DragScrollArea({ children }: { children: React.ReactNode }) {
  const containerRef = useRef<HTMLDivElement>(null);
  const dragRef = useRef({
    active: false,
    moved: false,
    scrollLeft: 0,
    startX: 0,
  });

  function endDrag(event: React.PointerEvent<HTMLDivElement>) {
    const container = containerRef.current;

    if (!container || !dragRef.current.active) return;

    dragRef.current.active = false;
    container.classList.remove("is-dragging");

    if (container.hasPointerCapture(event.pointerId)) {
      container.releasePointerCapture(event.pointerId);
    }
  }

  return (
    <div
      className="padel-bracket-scroll overflow-x-auto pb-4"
      onClickCapture={(event) => {
        if (!dragRef.current.moved) return;

        event.preventDefault();
        event.stopPropagation();
        dragRef.current.moved = false;
      }}
      onPointerCancel={endDrag}
      onPointerDown={(event) => {
        if (event.button !== 0) return;
        // On touch/pen let the browser scroll natively (with momentum); the
        // manual drag is a mouse-only affordance.
        if (event.pointerType !== "mouse") return;

        const target = event.target as HTMLElement;
        const interactive = target.closest(
          "a,button,input,select,textarea,[role='button']",
        );

        if (interactive) return;

        const container = containerRef.current;

        if (!container) return;

        dragRef.current = {
          active: true,
          moved: false,
          scrollLeft: container.scrollLeft,
          startX: event.clientX,
        };
        container.classList.add("is-dragging");
        container.setPointerCapture(event.pointerId);
      }}
      onPointerLeave={endDrag}
      onPointerMove={(event) => {
        const container = containerRef.current;

        if (!container || !dragRef.current.active) return;

        const deltaX = event.clientX - dragRef.current.startX;

        if (Math.abs(deltaX) > 4) dragRef.current.moved = true;

        container.scrollLeft = dragRef.current.scrollLeft - deltaX;
      }}
      onPointerUp={endDrag}
      ref={containerRef}
    >
      {children}
    </div>
  );
}

function BracketView({
  categoryId,
  draw,
  onSelectWinner,
  readOnly = false,
  schedule,
  title,
}: {
  categoryId: string;
  draw: Draw;
  onSelectWinner?: (categoryId: string, match: Match, team: Team) => void;
  readOnly?: boolean;
  schedule: Record<string, ScheduleAssignment>;
  title: string;
}) {
  const isConsolation = draw.kind === "consolation";
  const maxRoundMatches = Math.max(
    0,
    ...draw.rounds.map((round) => round.matches.length),
  );
  const densityClass =
    maxRoundMatches >= 16
      ? "draw-board-large"
      : maxRoundMatches >= 8
        ? "draw-board-medium"
        : "";

  return (
    <section className={`draw-board ${densityClass}`}>
      {!readOnly ? (
        <div className="draw-board-head">
          <div className="flex items-center gap-2.5">
            <span
              className={`draw-board-icon ${
                isConsolation
                  ? "draw-board-icon-blue"
                  : "draw-board-icon-green"
              }`}
            >
              {isConsolation ? (
                <Users className="h-5 w-5" />
              ) : (
                <Trophy className="h-5 w-5" />
              )}
            </span>
            <div>
              <p className="wc-kicker">
                {isConsolation ? "Consolacion" : "Principal"}
              </p>
              <h2 className="wc-section-title">{title}</h2>
              <p className="draw-board-copy mt-0.5 text-sm font-medium text-[var(--ink-soft)]">
                {draw.rounds.length
                  ? "Selecciona ganadores para avanzar"
                  : "Pendiente"}
              </p>
            </div>
          </div>
        </div>
      ) : null}

      {draw.rounds.length ? (
        <DragScrollArea>
          <div className="padel-bracket-frame">
            {readOnly ? (
              <h2 className="draw-board-inline-title">{title}</h2>
            ) : null}
            <div className="padel-bracket">
              {draw.rounds.map((round, roundIndex) => {
                const isFinal = roundIndex === draw.rounds.length - 1;

                return (
                  <div className="padel-round-column" key={round.id}>
                    <div
                      className={`wc-round-label mb-3 ${
                        isFinal
                          ? "wc-round-label-final"
                          : isConsolation
                            ? "wc-round-label-blue"
                            : ""
                      }`}
                    >
                      {round.name}
                    </div>
                    <div
                      className={`padel-round ${isFinal ? "padel-round--final" : ""}`}
                    >
                      {round.matches.map((match) => (
                        <div className="padel-match-cell" key={match.id}>
                          <MatchCard
                            categoryId={categoryId}
                            isFinal={isFinal}
                            match={match}
                            onSelectWinner={onSelectWinner}
                            readOnly={readOnly}
                            schedule={schedule[matchScheduleKey(categoryId, match)]}
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </DragScrollArea>
      ) : (
        <div className="wc-empty-inline m-4">
          La consolacion aparece cuando marques perdedores de primera ronda.
        </div>
      )}
    </section>
  );
}

function MatchCard({
  categoryId,
  isFinal = false,
  match,
  onSelectWinner,
  readOnly,
  schedule,
}: {
  categoryId: string;
  isFinal?: boolean;
  match: Match;
  onSelectWinner?: (categoryId: string, match: Match, team: Team) => void;
  readOnly: boolean;
  schedule?: ScheduleAssignment;
}) {
  const playable = isPlayableMatch(match);
  const champion = isFinal && Boolean(match.winner);
  // Conflicts are internal-only: never surface them on the public (read-only) view.
  const showConflict = Boolean(schedule?.conflict) && !readOnly;

  return (
    <article
      className={`wc-match ${
        champion
          ? "wc-match-champion"
          : showConflict
            ? "wc-match-conflict"
            : ""
      }`}
    >
      <div className="mb-2 flex items-start justify-between gap-2">
        <p className="wc-match-when min-w-0">
          {schedule ? (
            <>
              {schedule.dayLabel} {schedule.time}
              <span className="wc-match-court"> / Pista {schedule.court}</span>
            </>
          ) : (
            <span className="text-[var(--ink-faint)]">Sin horario</span>
          )}
        </p>
        <span className="wc-match-num">{match.id}</span>
      </div>
      <TeamButton
        active={match.winner?.id === match.sideA?.id}
        champion={champion}
        disabled={
          readOnly ||
          !playable ||
          !match.sideA ||
          Boolean(match.sideA.isPlaceholder)
        }
        onClick={() =>
          match.sideA && onSelectWinner?.(categoryId, match, match.sideA)
        }
        team={match.sideA}
      />
      <TeamButton
        active={match.winner?.id === match.sideB?.id}
        champion={champion}
        disabled={
          readOnly ||
          !playable ||
          !match.sideB ||
          Boolean(match.sideB.isPlaceholder)
        }
        onClick={() =>
          match.sideB && onSelectWinner?.(categoryId, match, match.sideB)
        }
        team={match.sideB}
      />
    </article>
  );
}

function TeamButton({
  active,
  champion = false,
  disabled,
  onClick,
  team,
}: {
  active: boolean;
  champion?: boolean;
  disabled: boolean;
  onClick: () => void;
  team: Team | null;
}) {
  return (
    <button
      className={`wc-team ${active ? "wc-team-won" : ""}`}
      disabled={disabled}
      onClick={onClick}
      type="button"
    >
      <span className="flex min-w-0 items-center gap-1.5">
        {active ? (
          champion ? (
            <Crown className="h-4 w-4 shrink-0 text-[var(--gold-500)]" />
          ) : (
            <CheckCircle2 className="h-4 w-4 shrink-0 text-[var(--pitch-600)]" />
          )
        ) : null}
        <span className="min-w-0 truncate">{team?.name ?? "Pendiente"}</span>
      </span>
      {team?.seed ? <span className="wc-seed">{team.seed}</span> : null}
    </button>
  );
}

function ScheduleEditor({
  categoryId,
  drawSet,
  manualScheduleOverrides,
  onClearSchedule,
  onRecalculateKeepingManual,
  onResetManualSchedule,
  onUpdateManualSchedule,
  scheduleConfig,
  schedule,
}: {
  categoryId: string;
  drawSet: CategoryDrawSet;
  manualScheduleOverrides: ManualScheduleMap;
  onClearSchedule: (categoryId: string) => void;
  onRecalculateKeepingManual: () => void;
  onResetManualSchedule: (scheduleKey: string) => void;
  onUpdateManualSchedule: (
    scheduleKey: string,
    override: ManualScheduleOverride,
  ) => void;
  scheduleConfig: TournamentScheduleConfig;
  schedule: Record<string, ScheduleAssignment>;
}) {
  const editableMatches = [
    ...drawSet.mainDraw.rounds.flatMap((round) =>
      round.matches.filter(matchNeedsSchedule).map((match) => ({
        drawName: "Principal",
        match,
        roundName: round.name,
      })),
    ),
    ...drawSet.consolationDraw.rounds.flatMap((round) =>
      round.matches.filter(matchNeedsSchedule).map((match) => ({
        drawName: "Consolacion",
        match,
        roundName: round.name,
      })),
    ),
  ];
  const [drafts, setDrafts] = useState<ManualScheduleMap>({});
  const [dirtyKeys, setDirtyKeys] = useState<Set<string>>(() => new Set());
  const pendingCount = dirtyKeys.size;

  function selectedValue(scheduleKey: string) {
    if (dirtyKeys.has(scheduleKey) && drafts[scheduleKey])
      return drafts[scheduleKey];

    return scheduleToManualOverride(
      manualScheduleOverrides[scheduleKey]
        ? manualOverrideToAssignment(manualScheduleOverrides[scheduleKey])
        : schedule[scheduleKey],
    );
  }

  function updateDraft(
    scheduleKey: string,
    patch: Partial<ManualScheduleOverride>,
  ) {
    const current = selectedValue(scheduleKey);
    const nextDay = patch.day ?? current.day;
    const times = timeOptionsForDay(nextDay, scheduleConfig);
    const nextTime =
      patch.time && times.includes(patch.time)
        ? patch.time
        : times.includes(current.time)
          ? current.time
          : times[0] ?? current.time;

    setDrafts((currentDrafts) => ({
      ...currentDrafts,
      [scheduleKey]: {
        court: patch.court ?? current.court,
        day: nextDay,
        time: nextTime,
      },
    }));
    setDirtyKeys((currentDirty) => new Set(currentDirty).add(scheduleKey));
  }

  function saveOne(scheduleKey: string) {
    const draft = drafts[scheduleKey];

    if (!draft) return;

    onUpdateManualSchedule(scheduleKey, draft);
    setDrafts((currentDrafts) => {
      const next = { ...currentDrafts };

      delete next[scheduleKey];

      return next;
    });
    setDirtyKeys((currentDirty) => {
      const next = new Set(currentDirty);

      next.delete(scheduleKey);

      return next;
    });
  }

  function lockCurrent(scheduleKey: string) {
    onUpdateManualSchedule(scheduleKey, selectedValue(scheduleKey));
  }

  function saveAll() {
    Array.from(dirtyKeys).forEach(saveOne);
  }

  function recalculateWithManual() {
    if (dirtyKeys.size) {
      saveAll();
      return;
    }

    onRecalculateKeepingManual();
  }

  function resetRow(scheduleKey: string) {
    if (dirtyKeys.has(scheduleKey)) {
      setDrafts((currentDrafts) => {
        const next = { ...currentDrafts };

        delete next[scheduleKey];

        return next;
      });
      setDirtyKeys((currentDirty) => {
        const next = new Set(currentDirty);

        next.delete(scheduleKey);

        return next;
      });
      return;
    }

    onResetManualSchedule(scheduleKey);
  }

  function resetAuto() {
    setDrafts({});
    setDirtyKeys(new Set());
    onClearSchedule(categoryId);
  }

  return (
    <section className="wc-pane">
      <div className="wc-pane-head">
        <div>
          <p className="wc-kicker">Horarios</p>
          <h2 className="wc-section-title">Bloqueos y recalculo</h2>
          <p className="wc-pane-copy">
            Los horarios bloqueados se fuerzan y el resto se recalcula alrededor.
          </p>
        </div>
        <div className="wc-pane-actions">
          <button
            className="wc-btn wc-btn-primary wc-btn-sm"
            onClick={recalculateWithManual}
            type="button"
          >
            <RotateCcw className="h-4 w-4" />
            {pendingCount
              ? "Bloquear cambios y recalcular"
              : "Recalcular con candados"}
            {pendingCount ? ` (${pendingCount})` : ""}
          </button>
          <button
            className="wc-btn wc-btn-ghost wc-btn-sm"
            onClick={resetAuto}
            type="button"
          >
            <RotateCcw className="h-4 w-4" />
            Restaurar horarios auto
          </button>
        </div>
      </div>

      <div className="wc-schedule-edit-list">
        {editableMatches.map(({ drawName, match, roundName }) => {
          const key = matchScheduleKey(categoryId, match);
          const manual = manualScheduleOverrides[key];
          const hasDraft = dirtyKeys.has(key);
          const current = selectedValue(key);
          const dayTimes = timeOptionsForDay(current.day, scheduleConfig);
          const times = dayTimes.includes(current.time)
            ? dayTimes
            : [current.time, ...dayTimes];
          const locked = Boolean(manual);

          return (
            <article
              className={`wc-schedule-edit-row ${
                schedule[key]?.conflict
                  ? "wc-schedule-edit-row-conflict"
                  : hasDraft
                    ? "wc-schedule-edit-row-draft"
                    : ""
              }`}
              key={key}
            >
              <div className="flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <p className="wc-row-label">
                    {drawName} / {roundName} / {match.id}
                  </p>
                  <p className="wc-schedule-match-label">
                    {matchTeamsLabel(match)}
                  </p>
                </div>
                <span
                  className={`wc-chip shrink-0 ${
                    hasDraft
                      ? "wc-chip-green"
                      : locked
                        ? "wc-chip-gold"
                        : "wc-chip-neutral"
                  }`}
                >
                  {hasDraft ? "Pendiente" : locked ? "Bloqueado" : "Auto"}
                </span>
              </div>

              <div className="wc-schedule-edit-controls">
                <select
                  className="wc-field h-9"
                  onChange={(event) =>
                    updateDraft(key, { day: event.target.value as DayKey })
                  }
                  value={current.day}
                >
                  {dayOptions.map((day) => (
                    <option key={day.key} value={day.key}>
                      {day.label}
                    </option>
                  ))}
                </select>
                <select
                  className="wc-field h-9 tabular"
                  onChange={(event) =>
                    updateDraft(key, { time: event.target.value })
                  }
                  value={current.time}
                >
                  {times.map((time) => (
                    <option key={time} value={time}>
                      {time}
                    </option>
                  ))}
                </select>
                <select
                  className="wc-field h-9 tabular"
                  onChange={(event) =>
                    updateDraft(key, {
                      court: Number.parseInt(event.target.value, 10),
                    })
                  }
                  value={current.court}
                >
                  {Array.from(
                    { length: courtCount },
                    (_, index) => index + 1,
                  ).map((court) => (
                    <option key={court} value={court}>
                      P{court}
                    </option>
                  ))}
                </select>
                <button
                  className={`wc-btn wc-btn-sm ${
                    locked && !hasDraft ? "wc-btn-ghost" : "wc-btn-primary"
                  }`}
                  disabled={!hasDraft && !locked && !schedule[key]}
                  onClick={() =>
                    locked && !hasDraft
                      ? resetRow(key)
                      : hasDraft
                        ? saveOne(key)
                        : lockCurrent(key)
                  }
                  type="button"
                >
                  {locked && !hasDraft ? (
                    <LockOpen className="h-4 w-4" />
                  ) : (
                    <LockKeyhole className="h-4 w-4" />
                  )}
                  {locked && !hasDraft ? "Soltar" : "Fijar"}
                </button>
                <button
                  className="wc-icon-btn h-9 w-9"
                  disabled={!hasDraft}
                  onClick={() => resetRow(key)}
                  title="Descartar cambio"
                  type="button"
                >
                  <RotateCcw className="h-4 w-4" />
                </button>
              </div>
            </article>
          );
        })}
      </div>
    </section>
  );
}

function PublicAdminPanel({
  copied,
  onCopy,
  onPublish,
  publicPath,
  publicUrl,
  published,
  remoteStatus,
  rows,
}: {
  copied: boolean;
  onCopy: () => void;
  onPublish: () => void;
  publicPath: string;
  publicUrl: string;
  published: boolean;
  remoteStatus: "idle" | "saving" | "ok" | "error";
  rows: ScheduleSummaryRow[];
}) {
  return (
    <div className="grid gap-5">
      <section className="wc-pane">
        <div className="wc-pane-head">
          <div>
            <p className="wc-kicker">Salida publica</p>
            <h2 className="wc-pane-title">Match centre listo para jugadores</h2>
            <p className="wc-pane-copy">
              Publica horarios, pistas, categorias y cuadros en una vista
              preparada para consulta rapida.
            </p>
          </div>
          <span className="wc-public-icon">
            <Globe2 className="h-6 w-6" />
          </span>
        </div>
      </section>

      <PublishPanel
        copied={copied}
        onCopy={onCopy}
        onPublish={onPublish}
        publicPath={publicPath}
        publicUrl={publicUrl}
        published={published}
        remoteStatus={remoteStatus}
      />
      <PublicSchedulePanel rows={rows} />
    </div>
  );
}

function PublishPanel({
  copied,
  onCopy,
  onPublish,
  publicPath,
  publicUrl,
  published,
  remoteStatus,
}: {
  copied: boolean;
  onCopy: () => void;
  onPublish: () => void;
  publicPath: string;
  publicUrl: string;
  published: boolean;
  remoteStatus: "idle" | "saving" | "ok" | "error";
}) {
  return (
    <section className="wc-publish-panel">
      <div className="min-w-0">
        <div>
          <div className="flex items-center gap-2">
            <p className="wc-kicker">Publicacion</p>
            {published ? (
              <span className="wc-chip wc-chip-green">
                <CheckCircle2 className="h-3.5 w-3.5" />
                En vivo
              </span>
            ) : null}
          </div>
          <h2 className="wc-section-title mt-1">
            {published ? "Cuadro publico listo" : "Publicar cuadro"}
          </h2>
          <p className="wc-pane-copy">
            Al publicar se guarda en la base de datos compartida, asi el enlace
            funciona en cualquier dispositivo y no se borra. Si aun no hay base
            de datos configurada, el enlace solo funciona en este navegador.
          </p>
          {remoteStatus === "saving" ? (
            <p className="mt-2 text-sm font-bold text-[var(--ink-soft)]">
              Guardando en la base de datos...
            </p>
          ) : remoteStatus === "ok" ? (
            <p className="mt-2 text-sm font-bold text-[var(--pitch-700)]">
              Guardado en la base de datos: visible en cualquier dispositivo.
            </p>
          ) : remoteStatus === "error" ? (
            <p className="mt-2 text-sm font-bold text-[var(--coral-600)]">
              Guardado solo en este navegador (base de datos no disponible).
            </p>
          ) : null}
        </div>
      </div>
      <div className="wc-publish-actions">
          <button className="wc-btn wc-btn-dark" onClick={onPublish} type="button">
            <Share2 className="h-4 w-4" />
            {published ? "Actualizar publico" : "Publicar"}
          </button>
          <button
            className="wc-btn wc-btn-ghost"
            disabled={!publicUrl}
            onClick={onCopy}
            type="button"
          >
            <Copy className="h-4 w-4" />
            {copied ? "Copiado" : "Copiar enlace"}
          </button>
          {publicPath ? (
            <a
              className="wc-btn wc-btn-ghost"
              href={publicPath}
              rel="noreferrer"
              target="_blank"
            >
              <ExternalLink className="h-4 w-4" />
              Ver publico
            </a>
          ) : null}
      </div>
    </section>
  );
}

function Notice({
  icon: Icon,
  text,
  title,
  tone,
}: {
  icon: typeof AlertTriangle;
  text: string;
  title: string;
  tone: "warning" | "success";
}) {
  return (
    <section
      className={`wc-notice ${
        tone === "warning" ? "wc-notice-warning" : "wc-notice-success"
      }`}
    >
      <Icon className="mt-0.5 h-5 w-5 shrink-0" />
      <div>
        <h2 className="font-display text-sm font-bold uppercase tracking-wide">
          {title}
        </h2>
        <p className="mt-0.5 text-sm font-semibold text-[var(--ink)]/80">{text}</p>
      </div>
    </section>
  );
}

function roundLabel(row: ScheduleSummaryRow) {
  return row.drawName === "Consolacion"
    ? `Consol. / ${row.roundName}`
    : row.roundName;
}

function groupRowsByDay(rows: ScheduleSummaryRow[]) {
  const groups: { dayLabel: string; rows: ScheduleSummaryRow[] }[] = [];

  rows.forEach((row) => {
    const last = groups[groups.length - 1];

    if (last && last.dayLabel === row.dayLabel) {
      last.rows.push(row);
    } else {
      groups.push({ dayLabel: row.dayLabel, rows: [row] });
    }
  });

  return groups;
}

function ScheduleRow({ row }: { row: ScheduleSummaryRow }) {
  const [teamA, teamB] = row.match.split(" vs ");

  return (
    <tr>
      <td className="wc-sched-hora">{row.time}</td>
      <td className="wc-sched-pista">{row.court ? `P${row.court}` : "-"}</td>
      <td>
        <span className="flex items-center gap-2">
          <span
            className="wc-dot"
            style={{ backgroundColor: categoryCrestColor(row.categoryId) }}
          />
          <span className="wc-sched-teams">
            {teamA}
            <em>vs</em>
            {teamB ?? "Pendiente"}
          </span>
        </span>
      </td>
      <td className="wc-sched-muted hidden md:table-cell">{row.categoryName}</td>
      <td className="wc-sched-muted hidden sm:table-cell">{roundLabel(row)}</td>
    </tr>
  );
}

function normalizeSearch(value: string) {
  return value
    .normalize("NFD")
    .replace(/\p{M}/gu, "")
    .toLowerCase();
}

function scheduleSearchText(row: ScheduleSummaryRow) {
  return normalizeSearch(
    [
      row.match,
      row.categoryName,
      row.drawName,
      row.roundName,
      row.dayLabel,
      row.time,
      row.court ? `p${row.court}` : "",
      row.court ? `pista ${row.court}` : "",
      row.scheduleKey,
    ].join(" "),
  );
}

function PublicSchedulePanel({ rows }: { rows: ScheduleSummaryRow[] }) {
  const [activeCategory, setActiveCategory] = useState("all");
  const [query, setQuery] = useState("");

  const categories = useMemo(() => {
    const seen = new Set<string>();
    const list: { id: string; name: string }[] = [];

    rows.forEach((row) => {
      if (!seen.has(row.categoryId)) {
        seen.add(row.categoryId);
        list.push({ id: row.categoryId, name: row.categoryName });
      }
    });

    return list;
  }, [rows]);

  const normalizedQuery = normalizeSearch(query.trim());
  const filtered = rows.filter((row) => {
    const matchesCategory =
      activeCategory === "all" || row.categoryId === activeCategory;
    const matchesQuery =
      !normalizedQuery || scheduleSearchText(row).includes(normalizedQuery);

    return matchesCategory && matchesQuery;
  });
  const dayGroups = groupRowsByDay(filtered);

  return (
    <section className="wc-public-schedule">
      <div className="wc-public-schedule-head">
        <div className="wc-public-schedule-title-row">
          <h2 className="wc-title text-2xl sm:text-3xl">Horarios del torneo</h2>
          <span className="wc-public-schedule-count">
            {filtered.length} de {rows.length} partidos
          </span>
        </div>

        <div className="wc-public-schedule-controls">
          <div className="wc-public-filter-row">
            <button
              className={`wc-pill ${activeCategory === "all" ? "wc-pill-active" : ""}`}
              onClick={() => setActiveCategory("all")}
              type="button"
            >
              Todas
            </button>
            {categories.map((category) => (
              <button
                className={`wc-pill ${
                  activeCategory === category.id ? "wc-pill-active" : ""
                }`}
                key={category.id}
                onClick={() => setActiveCategory(category.id)}
                type="button"
              >
                <span
                  className="wc-crest"
                  style={{ backgroundColor: categoryCrestColor(category.id) }}
                />
                {category.name}
              </button>
            ))}
          </div>

          <div className="wc-public-search">
            <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-[var(--ink-faint)]" />
            <label className="sr-only" htmlFor="public-schedule-search">
              Buscar horarios
            </label>
            <input
              id="public-schedule-search"
              className="wc-field w-full pl-9 pr-11"
              onChange={(event) => setQuery(event.target.value)}
              placeholder="Buscar jugador, pista u hora"
              type="search"
              value={query}
            />
            {query ? (
              <button
                className="absolute right-2 top-1/2 flex h-8 w-8 -translate-y-1/2 items-center justify-center rounded-full text-[var(--ink-faint)] transition hover:bg-[var(--surface-2)] hover:text-[var(--ink)]"
                onClick={() => setQuery("")}
                title="Limpiar busqueda"
                type="button"
              >
                <X className="h-4 w-4" />
              </button>
            ) : null}
          </div>
        </div>
      </div>

      {filtered.length ? (
        dayGroups.map((group) => (
          <div className="flex flex-col gap-2.5" key={group.dayLabel}>
            <div className="wc-day-head">
              <h3 className="wc-day-title">{group.dayLabel}</h3>
              <span className="text-xs font-semibold uppercase text-[var(--ink-faint)]">
                {group.rows.length} partidos
              </span>
              <span className="wc-day-rule" />
            </div>
            <div className="overflow-x-auto">
              <table className="wc-schedule">
                <thead>
                  <tr>
                    <th>Hora</th>
                    <th>Pista</th>
                    <th>Enfrentamiento</th>
                    <th className="hidden md:table-cell">Categoria</th>
                    <th className="hidden sm:table-cell">Ronda</th>
                  </tr>
                </thead>
                <tbody>
                  {group.rows.map((row) => (
                    <ScheduleRow key={row.scheduleKey} row={row} />
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        ))
      ) : (
        <div className="wc-empty-inline p-6">
          {rows.length
            ? "No hay partidos que coincidan con el filtro."
            : "Horarios pendientes."}
        </div>
      )}
    </section>
  );
}

export function TournamentPublicPage({ slug }: { slug: string }) {
  const [tournament, setTournament] = useState<PublishedTournament | null>(
    null,
  );
  const [activePublicCategoryId, setActivePublicCategoryId] = useState("");
  const [activePublicView, setActivePublicView] = useState<
    "schedule" | "draws"
  >("schedule");

  useEffect(() => {
    // `undefined` is a sentinel that can never equal a real payload, so the
    // first poll always applies (even when the tournament is unpublished and
    // reads back as null).
    let lastSignature: string | null | undefined;
    let cancelled = false;

    const localRaw = () =>
      typeof window === "undefined"
        ? null
        : window.localStorage.getItem(publicStorageKey(slug));

    const apply = (
      signature: string | null,
      next: PublishedTournament | null,
    ) => {
      // Skip the state update (and the full re-render it triggers) when the
      // payload has not changed since the last poll.
      if (cancelled || signature === lastSignature) return;

      lastSignature = signature;
      setTournament(next);
    };

    const read = async () => {
      try {
        const response = await fetch(`/api/public/${slug}`, {
          cache: "no-store",
        });
        const bodyText = await response.text();
        const data = JSON.parse(bodyText) as {
          configured: boolean;
          tournament: PublishedTournament | null;
        };

        if (data.configured) {
          // The shared store is the source of truth once it is set up.
          apply(bodyText, data.tournament);
          return;
        }
      } catch {
        // Network/API problem: fall back to this browser's local copy below.
      }

      const raw = localRaw();
      apply(raw, safeJsonParse<PublishedTournament>(raw));
    };

    read();
    const interval = window.setInterval(read, 5000);

    return () => {
      cancelled = true;
      window.clearInterval(interval);
    };
  }, [slug]);

  const drawSets = useMemo(
    () =>
      tournament
        ? buildDrawSets(
            tournament.categories,
            tournament.mainSelectionsByCategory,
            tournament.consolationSelectionsByCategory,
          )
        : [],
    [tournament],
  );
  const schedule = useMemo(
    () =>
      tournament
        ? buildGlobalSchedule(
            drawSets,
            tournament.manualScheduleOverrides,
            tournament.scheduleConfig ?? defaultScheduleConfig,
          )
        : { assignments: {}, conflicts: 0, saturdayCount: 0, total: 0 },
    [drawSets, tournament],
  );
  const publicSummaryRows = useMemo(
    () =>
      tournament
        ? buildScheduleSummaryRows(
            tournament.categories,
            drawSets,
            schedule.assignments,
          )
        : [],
    [drawSets, schedule.assignments, tournament],
  );
  const activeDrawSet =
    drawSets.find((drawSet) => drawSet.categoryId === activePublicCategoryId) ??
    drawSets[0] ??
    null;
  const activePublicId = activeDrawSet?.categoryId ?? "";

  if (!tournament) {
    return (
      <main className="wc-app grid min-h-screen place-items-center p-6">
        <section className="wc-not-published">
          <span className="rg-logo-mark mx-auto">
            <Trophy className="h-7 w-7" />
          </span>
          <h1>Cuadro no publicado</h1>
          <p>
            Publicalo desde el panel de administracion.
          </p>
        </section>
      </main>
    );
  }

  return (
    <main className="wc-app wc-public-app min-h-screen text-[var(--ink)]">
      <header className="wc-topbar">
        <div className="wc-topbar-inner">
          <div className="wc-public-brand">
            <span className="rg-logo-mark">
              <Trophy className="h-6 w-6" />
            </span>
            <div className="min-w-0">
              <p className="wc-kicker">Match centre</p>
              <h1 className="wc-brand-title truncate">
                {tournament.name}
              </h1>
            </div>
          </div>
        </div>
      </header>
      <div
        className={`wc-shell ${
          activePublicView === "draws" ? "wc-shell-draws" : ""
        }`}
      >
        <section className="wc-public-view-switch">
            {[
              { id: "schedule" as const, label: "Horarios" },
              { id: "draws" as const, label: "Cuadros" },
            ].map((tab) => (
              <button
                className={`wc-public-view-tab ${
                  activePublicView === tab.id ? "wc-public-view-tab-active" : ""
                }`}
                key={tab.id}
                onClick={() => setActivePublicView(tab.id)}
                type="button"
              >
                {tab.label}
              </button>
            ))}
        </section>

        {activePublicView === "schedule" ? (
          <PublicSchedulePanel rows={publicSummaryRows} />
        ) : null}

        {activePublicView === "draws" ? (
          <>
            <section className="wc-category-strip">
              <div className="wc-category-strip-inner">
                {drawSets.map((drawSet) => (
                  <button
                    className={`wc-pill ${
                      drawSet.categoryId === activePublicId ? "wc-pill-active" : ""
                    }`}
                    key={drawSet.categoryId}
                    onClick={() =>
                      setActivePublicCategoryId(drawSet.categoryId)
                    }
                    type="button"
                  >
                    <span
                      className="wc-crest"
                      style={{
                        backgroundColor: categoryCrestColor(drawSet.categoryId),
                      }}
                    />
                    {drawSet.categoryName}
                  </button>
                ))}
              </div>
            </section>

            {activeDrawSet ? (
              <section className="wc-bracket-field wc-public-draw-field">
                <BracketView
                  categoryId={activeDrawSet.categoryId}
                  draw={activeDrawSet.mainDraw}
                  readOnly
                  schedule={schedule.assignments}
                  title="Cuadro principal"
                />
                <BracketView
                  categoryId={activeDrawSet.categoryId}
                  draw={activeDrawSet.consolationDraw}
                  readOnly
                  schedule={schedule.assignments}
                  title="Consolacion"
                />
              </section>
            ) : null}
          </>
        ) : null}
      </div>
    </main>
  );
}
