export type MapRegionElement = SVGPathElement;

export type MapRegionEvent = {
  id: string;
  name: string;
  element: MapRegionElement;
};

type RetrieveMultipleResponse = {
  entities: Array<Record<string, unknown>>;
};

type DataverseWebApi = {
  retrieveMultipleRecords: (
    entityLogicalName: string,
    options?: string,
    maxPageSize?: number,
  ) => Promise<RetrieveMultipleResponse>;
};

type DataverseContext = {
  webAPI: DataverseWebApi;
};

type CountryMetrics = {
  fairs: number;
  leads: number;
};

export type BindInteractiveWorldMapOptions = {
  context: DataverseContext;
  hoverFill?: string;
  activeFill?: string;
  hoverScale?: number;
  clickScale?: number;
  animationMs?: number;
  tooltipOffset?: number;
  regionSelector?: string;
  countryFieldName?: string;
  fairsEntityName?: string;
  leadsEntityName?: string;
  onClick?: (region: MapRegionEvent, metrics: CountryMetrics) => void;
};

const STYLE_ID = "interactive-world-map-styles";
const TOOLTIP_CLASS = "interactive-world-map-tooltip";

function ensureStyles(
  options: Required<
    Pick<
      BindInteractiveWorldMapOptions,
      "hoverFill" | "activeFill" | "hoverScale" | "clickScale" | "animationMs"
    >
  >,
): void {
  if (document.getElementById(STYLE_ID)) {
    return;
  }

  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = `
.interactive-world-map {
  overflow: visible;
}

.interactive-world-map [data-map-region] {
  cursor: pointer;
  transform-box: fill-box;
  transform-origin: center;
  transition:
    fill ${options.animationMs}ms ease,
    stroke ${options.animationMs}ms ease,
    stroke-width ${options.animationMs}ms ease,
    transform ${options.animationMs}ms ease,
    opacity ${options.animationMs}ms ease;
}

.interactive-world-map [data-map-region].is-hovered {
  fill: ${options.hoverFill};
  stroke: #0b3c5d;
  stroke-width: 0.9;
  transform: scale(${options.hoverScale});
}

.interactive-world-map [data-map-region].is-active {
  fill: ${options.activeFill};
  stroke: #8d3b00;
  stroke-width: 1;
  transform: scale(${options.clickScale});
}

.interactive-world-map [data-map-region].is-clicked {
  animation: map-click-pulse ${options.animationMs}ms ease;
}

.${TOOLTIP_CLASS} {
  position: fixed;
  z-index: 9999;
  min-width: 180px;
  padding: 10px 12px;
  border: 1px solid #d0d7de;
  border-radius: 10px;
  background: rgba(255, 255, 255, 0.98);
  box-shadow: 0 10px 24px rgba(16, 24, 40, 0.12);
  color: #102a43;
  font-family: "Segoe UI", Arial, sans-serif;
  font-size: 13px;
  line-height: 1.45;
  pointer-events: none;
  opacity: 0;
  transform: translateY(4px);
  transition: opacity 120ms ease, transform 120ms ease;
}

.${TOOLTIP_CLASS}.is-visible {
  opacity: 1;
  transform: translateY(0);
}

.${TOOLTIP_CLASS} strong {
  font-weight: 700;
}

.${TOOLTIP_CLASS} .country-name {
  margin-bottom: 6px;
  font-weight: 700;
}

@keyframes map-click-pulse {
  0% {
    transform: scale(1);
  }

  50% {
    transform: scale(0.97);
  }

  100% {
    transform: scale(1);
  }
}
`;

  document.head.appendChild(style);
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function getRegionMeta(element: MapRegionElement): MapRegionEvent {
  return {
    id: element.id || "",
    name:
      element.getAttribute("name") ||
      element.getAttribute("class") ||
      element.id ||
      "",
    element,
  };
}

function createTooltip(): HTMLDivElement {
  const tooltip = document.createElement("div");
  tooltip.className = TOOLTIP_CLASS;
  document.body.appendChild(tooltip);
  return tooltip;
}

function setTooltipContent(
  tooltip: HTMLDivElement,
  countryName: string,
  metrics: CountryMetrics | null,
): void {
  const fairs = metrics ? metrics.fairs : 0;
  const leads = metrics ? metrics.leads : 0;

  tooltip.innerHTML = `
    <div class="country-name">${countryName}</div>
    <div><strong>No. of fairs:</strong> ${fairs}</div>
    <div><strong>No. of leads:</strong> ${leads}</div>
  `;
}

function positionTooltip(
  tooltip: HTMLDivElement,
  x: number,
  y: number,
  offset: number,
): void {
  tooltip.style.left = `${x + offset}px`;
  tooltip.style.top = `${y + offset}px`;
}

async function retrieveCountryCount(
  context: DataverseContext,
  entityLogicalName: string,
  countryFieldName: string,
  countryName: string,
): Promise<number> {
  const fetchXml = `
    <fetch aggregate="true">
      <entity name="${entityLogicalName}">
        <attribute name="${countryFieldName}" alias="recordcount" aggregate="countcolumn" />
        <filter>
          <condition attribute="${countryFieldName}" operator="eq" value="${escapeXml(countryName)}" />
        </filter>
      </entity>
    </fetch>
  `.trim();

  const response = await context.webAPI.retrieveMultipleRecords(
    entityLogicalName,
    `?fetchXml=${encodeURIComponent(fetchXml)}`,
  );

  const rawCount = response.entities[0]?.recordcount;
  const parsed =
    typeof rawCount === "number"
      ? rawCount
      : Number.parseInt(String(rawCount ?? "0"), 10);

  return Number.isNaN(parsed) ? 0 : parsed;
}

async function loadCountryMetrics(
  context: DataverseContext,
  countryName: string,
  countryFieldName: string,
  fairsEntityName: string,
  leadsEntityName: string,
): Promise<CountryMetrics> {
  const [fairs, leads] = await Promise.all([
    retrieveCountryCount(context, fairsEntityName, countryFieldName, countryName),
    retrieveCountryCount(context, leadsEntityName, countryFieldName, countryName),
  ]);

  return { fairs, leads };
}

export function bindInteractiveWorldMap(
  svgRoot: SVGSVGElement,
  options: BindInteractiveWorldMapOptions,
): () => void {
  const resolved = {
    hoverFill: options.hoverFill ?? "#8ecae6",
    activeFill: options.activeFill ?? "#ffb703",
    hoverScale: options.hoverScale ?? 1.015,
    clickScale: options.clickScale ?? 0.985,
    animationMs: options.animationMs ?? 180,
    tooltipOffset: options.tooltipOffset ?? 14,
    regionSelector: options.regionSelector ?? "path[id], path[name], path[class]",
    countryFieldName: options.countryFieldName ?? "new_country",
    fairsEntityName: options.fairsEntityName ?? "new_fairs",
    leadsEntityName: options.leadsEntityName ?? "new_leads",
    context: options.context,
    onClick: options.onClick,
  };

  ensureStyles(resolved);
  svgRoot.classList.add("interactive-world-map");

  const tooltip = createTooltip();
  const metricsCache = new Map<string, CountryMetrics>();
  const regions = Array.from(
    svgRoot.querySelectorAll<MapRegionElement>(resolved.regionSelector),
  );

  let hovered: MapRegionElement | null = null;
  let active: MapRegionElement | null = null;
  let activeCountry = "";

  const showTooltip = (countryName: string, metrics: CountryMetrics | null): void => {
    setTooltipContent(tooltip, countryName, metrics);
    tooltip.classList.add("is-visible");
  };

  const hideTooltip = (): void => {
    tooltip.classList.remove("is-visible");
  };

  const clearHover = (): void => {
    if (hovered && hovered !== active) {
      hovered.classList.remove("is-hovered");
    }

    hovered = null;
  };

  const moveTooltipFromEvent = (event: MouseEvent): void => {
    positionTooltip(
      tooltip,
      event.clientX,
      event.clientY,
      resolved.tooltipOffset,
    );
  };

  const handleEnter = (element: MapRegionElement, event: MouseEvent): void => {
    if (hovered && hovered !== active) {
      hovered.classList.remove("is-hovered");
    }

    hovered = element;

    if (element !== active) {
      element.classList.add("is-hovered");
    }

    if (activeCountry) {
      moveTooltipFromEvent(event);
      tooltip.classList.add("is-visible");
    }
  };

  const handleLeave = (element: MapRegionElement): void => {
    if (element !== active) {
      element.classList.remove("is-hovered");
    }

    if (hovered === element) {
      hovered = null;
    }

    if (!activeCountry) {
      hideTooltip();
    }
  };

  const handleClick = async (
    element: MapRegionElement,
    event: MouseEvent,
  ): Promise<void> => {
    const region = getRegionMeta(element);
    const countryName = region.name;

    if (!countryName) {
      return;
    }

    if (active && active !== element) {
      active.classList.remove("is-active");
    }

    if (active === element) {
      element.classList.remove("is-active");
      active = null;
      activeCountry = "";
      hideTooltip();
      return;
    }

    clearHover();
    element.classList.add("is-active");
    element.classList.remove("is-clicked");
    element.getBBox();
    element.classList.add("is-clicked");

    window.setTimeout(() => {
      element.classList.remove("is-clicked");
    }, resolved.animationMs);

    active = element;
    activeCountry = countryName;

    moveTooltipFromEvent(event);

    const cachedMetrics = metricsCache.get(countryName) ?? null;
    showTooltip(countryName, cachedMetrics);

    try {
      const metrics =
        cachedMetrics ??
        (await loadCountryMetrics(
          resolved.context,
          countryName,
          resolved.countryFieldName,
          resolved.fairsEntityName,
          resolved.leadsEntityName,
        ));

      metricsCache.set(countryName, metrics);
      showTooltip(countryName, metrics);
      resolved.onClick?.(region, metrics);
    } catch {
      const fallbackMetrics = { fairs: 0, leads: 0 };
      showTooltip(countryName, fallbackMetrics);
      resolved.onClick?.(region, fallbackMetrics);
    }
  };

  const cleanup: Array<() => void> = [];

  regions.forEach((region) => {
    region.dataset.mapRegion = "true";

    const onMouseEnter = (event: MouseEvent): void =>
      handleEnter(region, event);
    const onMouseMove = (event: MouseEvent): void => {
      if (tooltip.classList.contains("is-visible")) {
        moveTooltipFromEvent(event);
      }
    };
    const onMouseLeave = (): void => handleLeave(region);
    const onClick = (event: MouseEvent): void => {
      void handleClick(region, event);
    };

    region.addEventListener("mouseenter", onMouseEnter);
    region.addEventListener("mousemove", onMouseMove);
    region.addEventListener("mouseleave", onMouseLeave);
    region.addEventListener("click", onClick);

    cleanup.push(() => {
      region.removeEventListener("mouseenter", onMouseEnter);
      region.removeEventListener("mousemove", onMouseMove);
      region.removeEventListener("mouseleave", onMouseLeave);
      region.removeEventListener("click", onClick);
      region.classList.remove("is-hovered", "is-active", "is-clicked");
      delete region.dataset.mapRegion;
    });
  });

  return () => {
    clearHover();

    if (active) {
      active.classList.remove("is-active");
      active = null;
    }

    activeCountry = "";
    hideTooltip();
    tooltip.remove();

    cleanup.forEach((dispose) => dispose());
    svgRoot.classList.remove("interactive-world-map");
  };
}
