<script setup>
import { computed, nextTick, onBeforeUnmount, onMounted, ref } from "vue";
import * as XLSX from "xlsx";
import L from "leaflet";
import "leaflet/dist/leaflet.css";

const mapElement = ref(null);
const deliveries = ref([]);
const places = ref([]);
const activeIndex = ref(0);
const isPlaying = ref(true);
const isLoading = ref(true);
const errorMessage = ref("");
const showBorder = ref(true);
const showFront = ref(false);
let map;
let timer;
let routeLayers = [];
let placeLayers = [];
let borderLayer;
let frontLayer;
const localFrontData = new Map();

const fallbackUkraineBorder = [
  [51.5, 23.6],
  [51.3, 24.1],
  [51.4, 25.0],
  [51.1, 26.4],
  [50.4, 27.2],
  [50.5, 28.7],
  [50.0, 30.0],
  [50.5, 31.8],
  [51.1, 32.5],
  [51.3, 33.8],
  [51.0, 35.0],
  [50.4, 36.5],
  [49.8, 37.8],
  [49.2, 39.1],
  [48.4, 39.8],
  [47.8, 39.6],
  [47.1, 38.2],
  [46.6, 36.8],
  [46.1, 35.1],
  [45.3, 34.4],
  [45.4, 33.0],
  [46.1, 31.8],
  [46.5, 30.7],
  [46.0, 29.6],
  [45.6, 29.4],
  [46.3, 28.2],
  [47.0, 27.2],
  [47.4, 26.1],
  [48.2, 24.9],
  [48.6, 24.0],
  [49.0, 23.5],
  [49.5, 22.8],
  [49.9, 22.2],
  [50.4, 22.7],
  [50.8, 23.2],
  [51.5, 23.6],
];

const frontSnapshots = [
  {
    until: "2022-05-01",
    label: "Frühjahr 2022 · Näherung",
    coordinates: [
      [51.1, 31.0],
      [50.8, 32.4],
      [50.2, 33.5],
      [49.5, 34.8],
      [48.8, 36.1],
      [47.7, 37.9],
      [46.7, 39.0],
    ],
  },
  {
    until: "2022-10-01",
    label: "Sommer 2022 · Näherung",
    coordinates: [
      [51.2, 31.7],
      [50.7, 33.1],
      [50.1, 34.7],
      [49.4, 36.0],
      [48.6, 37.1],
      [47.8, 38.2],
      [46.8, 39.1],
    ],
  },
  {
    until: "2023-06-01",
    label: "Winter 2022/23 · Näherung",
    coordinates: [
      [51.0, 32.8],
      [50.2, 34.7],
      [49.4, 36.0],
      [48.6, 37.3],
      [47.7, 38.1],
      [46.8, 39.0],
    ],
  },
  {
    until: "2024-01-01",
    label: "Jahr 2023 · Näherung",
    coordinates: [
      [50.5, 34.3],
      [49.7, 35.6],
      [49.0, 36.8],
      [48.1, 37.7],
      [47.4, 38.1],
      [46.8, 39.0],
    ],
  },
  {
    until: "2100-01-01",
    label: "Ab 2024 · Näherung",
    coordinates: [
      [50.3, 34.5],
      [49.5, 35.7],
      [48.8, 36.7],
      [48.0, 37.5],
      [47.2, 38.0],
      [46.7, 39.0],
    ],
  },
];

const warContextByTour = {
  1: "Während dieser ersten Fahrt zogen sich russische Truppen aus dem Raum um Kyjiw zurück. Der Schwerpunkt des Krieges verlagerte sich im April 2022 in den Donbas.",
  2: "Im Sommer 2022 konzentrierten sich die besonders intensiven Kämpfe auf den Donbas. Die Ukraine bereitete zugleich ihre Gegenoffensive im Süden vor.",
  3: "Im Dezember 2022 verlief die Front vor allem im Donbas, unter anderem im Raum Bachmut. Russland setzte dort seine Offensive fort und griff wiederholt die ukrainische Energieinfrastruktur an.",
  4: "Zu Beginn des Jahres 2023 stand die russische Winteroffensive im Mittelpunkt. Besonders heftig waren die Kämpfe um Bachmut.",
  5: "Im Frühjahr 2023 dauerten die Kämpfe um Bachmut an. Die Ukraine bereitete parallel ihre Gegenoffensive vor.",
  6: "Im Spätsommer 2023 setzte die Ukraine ihre Gegenoffensive fort, stieß aber nur langsam vor. Russland hatte entlang der Front umfangreiche Verteidigungsstellungen angelegt.",
  7: "Anfang 2024 blieb die Front weitgehend statisch, während Russland seinen Druck im Osten erhöhte. Nach dem Fall von Awdijiwka rückten russische Truppen westlich der Stadt weiter vor.",
  8: "Im Herbst 2024 konzentrierten sich die Kämpfe weiterhin auf den Donbas und die Gebiete um Charkiw. Russland erzielte örtliche Geländegewinne.",
  9: "Im Frühjahr 2025 setzte Russland seine Angriffe im Osten und Süden fort. Raketen- und Drohnenangriffe auf ukrainische Städte und Infrastruktur blieben prägend.",
  10: "Im Sommer 2025 blieb der Abnutzungskrieg entlang der langen Front bestimmend. Lokale Vorstöße und Gegenangriffe wechselten sich ab.",
  11: "Im Herbst 2025 hielten die schweren Kämpfe im Osten und Süden an. Beide Seiten setzten verstärkt Drohnen, Artillerie und weitreichende Angriffe ein.",
  12: "Im Frühjahr 2026 setzte sich der Stellungskrieg mit örtlichen Angriffen und Gegenangriffen fort. Auch die gegenseitigen Drohnen- und Luftangriffe gingen weiter.",
  13: "Im Sommer 2026 blieb die Lage von langsamen, lokal begrenzten Frontbewegungen und hoher Angriffsdichte geprägt. Eine eindeutige strategische Entscheidung war nicht absehbar.",
};

const activeDelivery = computed(
  () => deliveries.value[activeIndex.value] || null,
);
const progress = computed(() =>
  deliveries.value.length
    ? ((activeIndex.value + 1) / deliveries.value.length) * 100
    : 0,
);
const cumulative = computed(() => {
  return deliveries.value.slice(0, activeIndex.value + 1).reduce(
    (totals, delivery) => ({
      vehicles: totals.vehicles + delivery.vehicles,
      people: totals.people + delivery.people,
      weight: totals.weight + delivery.weight,
      duration: totals.duration + delivery.duration,
      distance: totals.distance + delivery.distance,
    }),
    { vehicles: 0, people: 0, weight: 0, duration: 0, distance: 0 },
  );
});
const totals = computed(() =>
  deliveries.value.reduce(
    (result, delivery) => ({
      vehicles: result.vehicles + delivery.vehicles,
      people: result.people + delivery.people,
      weight: result.weight + delivery.weight,
      duration: result.duration + delivery.duration,
      distance: result.distance + delivery.distance,
    }),
    { vehicles: 0, people: 0, weight: 0, duration: 0, distance: 0 },
  ),
);

const chartMetrics = computed(() => [
  {
    key: "vehicles",
    label: "Fahrzeuge",
    value: cumulative.value.vehicles,
    unit: "",
    color: "coral",
  },
  {
    key: "people",
    label: "Personen",
    value: cumulative.value.people,
    unit: "",
    color: "blue",
  },
  {
    key: "weight",
    label: "Gewicht",
    value: cumulative.value.weight,
    unit: " Tonnen",
    color: "gold",
  },
  {
    key: "duration",
    label: "Dauer",
    value: cumulative.value.duration,
    unit: " Tage",
    color: "mint",
  },
  {
    key: "distance",
    label: "Distanz",
    value: cumulative.value.distance,
    unit: " km",
    color: "ink",
  },
]);

const formatDate = (serial) => {
  const date = XLSX.SSF.parse_date_code(serial);
  return date
    ? new Intl.DateTimeFormat("de-CH", {
        day: "2-digit",
        month: "short",
        year: "numeric",
      }).format(new Date(date.y, date.m - 1, date.d))
    : "-";
};
const formatNumber = (value) => new Intl.NumberFormat("de-CH").format(value);
const routeOpacity = (index) =>
  index === activeIndex.value ? 1 : index < activeIndex.value ? 0.6 : 0.12;

const loadData = async () => {
  try {
    const readWorkbook = async (file) =>
      XLSX.read(await (await fetch(`/data/${file}`)).arrayBuffer(), {
        type: "array",
      });
    const [deliveryBook, routeBook, placeBook] = await Promise.all(
      ["lieferung.xlsx", "strecke.xlsx", "orte.xlsx"].map(readWorkbook),
    );
    const deliveryRows = XLSX.utils.sheet_to_json(
      deliveryBook.Sheets[deliveryBook.SheetNames[0]],
      { defval: null },
    );
    const routeRows = XLSX.utils.sheet_to_json(
      routeBook.Sheets[routeBook.SheetNames[0]],
      { defval: null },
    );
    const placeRows = XLSX.utils.sheet_to_json(
      placeBook.Sheets[placeBook.SheetNames[0]],
      { defval: null },
    );
    const coordinateByName = new Map(
      placeRows.map((place) => [
        String(place.ort).trim().toLowerCase(),
        [place.lat, place.lon],
      ]),
    );

    places.value = placeRows.filter((place) =>
      coordinateByName.has(String(place.ort).trim().toLowerCase()),
    );
    deliveries.value = deliveryRows.map((row) => {
      const routeRow = routeRows.find((route) => route.tour === row.tour);
      const route = Object.keys(routeRow || {})
        .filter((key) => key.startsWith("ort_"))
        .map((key) => String(routeRow[key] || "").trim())
        .filter(Boolean);
      return {
        id: row.tour,
        startSerial: row.datum_s,
        start: formatDate(row.datum_s),
        end: formatDate(row.datum_e),
        vehicles: row.fz || 0,
        people: row.pers || 0,
        weight: row.gew || 0,
        duration: row.dauer || 0,
        distance: row.dist || 0,
        note: row.note,
        route,
        coordinates: route
          .map((place) => coordinateByName.get(place.toLowerCase()))
          .filter(Boolean),
      };
    });
    await Promise.all(
      deliveries.value.map(async (delivery) => {
        const date = XLSX.SSF.parse_date_code(delivery.startSerial);
        const dateKey = `${date.y}-${String(date.m).padStart(2, "0")}-${String(date.d).padStart(2, "0")}`;
        const response = await fetch(`/data/front/${dateKey}_daily.geo.json`);
        if (response.ok) localFrontData.set(delivery.id, await response.json());
      }),
    );
    isLoading.value = false;
    await nextTick();
    initMap();
    startPlayback();
  } catch (error) {
    isLoading.value = false;
    errorMessage.value = `Die Daten konnten nicht geladen werden: ${error.message}`;
  }
};

const initMap = () => {
  map = L.map(mapElement.value, {
    zoomControl: false,
    attributionControl: true,
  }).setView([49.2, 20.3], 5);
  L.control.zoom({ position: "bottomright" }).addTo(map);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 25,
    attribution: "&copy; OpenStreetMap-Mitwirkende",
  }).addTo(map);
  loadBorder();
  drawMap();
  const allDeliveryCoordinates = deliveries.value.flatMap(
    (delivery) => delivery.coordinates,
  );
  if (allDeliveryCoordinates.length) {
    map.fitBounds(L.latLngBounds(allDeliveryCoordinates), {
      padding: [0, 0],
      maxZoom: 20,
      animate: false,
    });
  }
};

const loadBorder = async () => {
  try {
    const response = await fetch(
      "https://raw.githubusercontent.com/johan/world.geo.json/master/countries/UKR.geo.json",
    );
    if (!response.ok) throw new Error("Ukraine boundary unavailable");
    borderLayer = L.geoJSON(await response.json(), {
      style: {
        color: "#1769aa",
        weight: 3,
        opacity: 0.95,
        fill: false,
        interactive: false,
      },
    }).addTo(map);
    borderLayer.bringToFront();
  } catch {
    borderLayer = L.polyline(fallbackUkraineBorder, {
      color: "#1769aa",
      weight: 3,
      opacity: 0.95,
      interactive: false,
    }).addTo(map);
  }
  updateOverlayVisibility();
};

const updateOverlayVisibility = () => {
  if (borderLayer)
    borderLayer.setStyle({ opacity: showBorder.value ? 0.95 : 0 });
  if (frontLayer)
    frontLayer.setStyle({
      opacity: showFront.value ? 0.95 : 0,
      fillOpacity: showFront.value ? 0.18 : 0,
    });
};

const drawFront = () => {
  if (frontLayer) frontLayer.remove();
  const geoJson = localFrontData.get(activeDelivery.value?.id);
  if (!geoJson) {
    return;
  }
  frontLayer = L.geoJSON(geoJson, {
    style: (feature) => ({
      color: feature.properties?.name === "russia" ? "#b11f36" : "#d17942",
      weight: 2,
      opacity: showFront.value ? 0.95 : 0,
      fillColor: feature.properties?.name === "russia" ? "#b11f36" : "#d17942",
      fillOpacity: showFront.value ? 0.18 : 0,
      dashArray: "7 6",
      interactive: false,
    }),
  }).addTo(map);
  frontLayer.bindTooltip(
    "Historische Gebietskontrolle · Frontverlauf als Grenze",
    { sticky: true },
  );
  frontLayer.bringToFront();
};

const drawMap = () => {
  routeLayers.forEach((layer) => layer.remove());
  placeLayers.forEach((layer) => layer.remove());
  drawFront();
  routeLayers = deliveries.value.map((delivery, index) =>
    L.polyline(delivery.coordinates, {
      color: index === activeIndex.value ? "#f2c230" : "#4f6f91",
      weight: index === activeIndex.value ? 4 : 2,
      opacity: routeOpacity(index),
      dashArray: index > activeIndex.value ? "4 8" : null,
    }).addTo(map),
  );
  placeLayers = places.value.map((place) => {
    const marker = L.circleMarker([place.lat, place.lon], {
      radius: place.ort === "Thun" ? 7 : 5,
      color: place.ort === "Thun" ? "#0057b7" : "#16324f",
      fillColor: "#ffffff",
      fillOpacity: 1,
      weight: 2,
    }).addTo(map);
    marker.bindTooltip(place.ort, { direction: "top", offset: [0, -5] });
    return marker;
  });
  updateOverlayVisibility();
};

const selectDelivery = (index) => {
  activeIndex.value = index;
  drawMap();
};
const toggleBorder = () => {
  showBorder.value = !showBorder.value;
  updateOverlayVisibility();
};
const toggleFront = () => {
  showFront.value = !showFront.value;
  updateOverlayVisibility();
};
const togglePlayback = () => {
  isPlaying.value ? stopPlayback() : startPlayback();
};
const startPlayback = () => {
  stopPlayback();
  isPlaying.value = true;
  timer = window.setInterval(() => {
    if (activeIndex.value < deliveries.value.length - 1)
      selectDelivery(activeIndex.value + 1);
    else stopPlayback();
  }, 5000);
};
const stopPlayback = () => {
  if (timer) window.clearInterval(timer);
  timer = null;
  isPlaying.value = false;
};
onMounted(loadData);
onBeforeUnmount(() => {
  stopPlayback();
  map?.remove();
});
</script>

<template>
  <main class="app-shell">
    <header class="topbar">
      <div class="brand">
        <span class="brand-mark">
          <img src="/data/mammutli.png" alt="Mammutli Logo" /></span
        ><span>Hilfslieferungen</span><small>UKRAINE · 2022 — 2026</small>
      </div>
      <div class="topbar-meta">
        <span class="live-dot"></span> <span>LIVE-CHRONIK</span
        ><strong>{{ deliveries.length || "—" }} LIEFERUNGEN</strong>
      </div>
    </header>
    <section class="intro">
      <div>
        <p class="eyebrow">Wegstrecke der Hilfslieferungen</p>
        <h1>Von Thun<br /><em>in die Ukraine.</em></h1>
      </div>
      <p class="intro-copy">
        Jemandem zu helfen verändert vielleicht nicht die Welt. Aber es kann die
        Welt für diese eine Person verändern. <br />(Helen Barry)
      </p>
    </section>
    <div v-if="errorMessage" class="status error">{{ errorMessage }}</div>
    <section v-else class="dashboard">
      <div class="map-panel panel">
        <div ref="mapElement" class="map"></div>
        <div v-if="isLoading" class="map-loading">
          Karte und Touren werden geladen …
        </div>
        <div class="map-overlay">
          <span class="route-key"
            ><i class="key-line active"></i> aktive Tour</span
          ><span class="route-key"
            ><i class="key-line"></i> gefahrene Strecke</span
          ><span class="route-key"
            ><i class="key-dot"></i> angefahrener Ort</span
          >
          <label class="layer-toggle"
            ><input
              type="checkbox"
              :checked="showBorder"
              @change="toggleBorder"
            />
            <i class="key-line border-key"></i> Ukraine-Grenze</label
          >
          <label class="layer-toggle"
            ><input
              type="checkbox"
              :checked="showFront"
              @change="toggleFront"
            />
            <i class="key-line front-key"></i> Historische Gebiete /
            Front</label
          >
        </div>
      </div>
      <aside class="side-panel">
        <div class="tour-heading">
          <div>
            <p class="eyebrow">Aktuelle Etappe</p>
            <h2>Tour {{ activeDelivery?.id || "—" }}</h2>
          </div>
          <span class="tour-count"
            >{{ String(activeIndex + 1).padStart(2, "0") }} /
            {{ String(deliveries.length).padStart(2, "0") }}</span
          >
        </div>
        <div v-if="activeDelivery" class="tour-detail">
          <div class="date-range">
            {{ activeDelivery.start }} <span>→</span> {{ activeDelivery.end }}
          </div>
          <div class="route-list">
            <span
              v-for="(place, index) in activeDelivery.route"
              :key="`${place}-${index}`"
              :class="{
                destination:
                  index > 0 && index < activeDelivery.route.length - 1,
              }"
              >{{ place }}</span
            >
          </div>
          <p v-if="activeDelivery.note" class="note">
            {{ activeDelivery.note }}
          </p>
          <p v-if="showFront" class="front-status">
            <span></span
            >{{
              warContextByTour[activeDelivery.id] ||
              "Während dieser Lieferung blieb die Frontlage in Bewegung."
            }}
          </p>
        </div>
        <div class="playback">
          <button
            class="play-button"
            type="button"
            @click="togglePlayback"
            :aria-label="isPlaying ? 'Pause' : 'Abspielen'"
          >
            {{ isPlaying ? "Ⅱ" : "▶" }}
          </button>
          <div class="progress-track">
            <span :style="{ width: `${progress}%` }"></span>
          </div>
          <span class="play-label">{{
            isPlaying ? "Ablauf läuft" : "Ablauf pausiert"
          }}</span>
        </div>
        <div class="tour-selector">
          <button
            v-for="(delivery, index) in deliveries"
            :key="delivery.id"
            type="button"
            :class="{ selected: index === activeIndex }"
            @click="selectDelivery(index)"
          >
            {{ delivery.id }}
          </button>
        </div>
      </aside>
    </section>
    <section class="metrics-section">
      <div class="section-heading">
        <div>
          <p class="eyebrow">Kumulierte Bilanz</p>
          <h2>Jede Fahrt zählt.</h2>
        </div>
        <span>Stand nach Tour {{ activeDelivery?.id || "—" }}</span>
      </div>
      <div class="chart">
        <div v-for="metric in chartMetrics" :key="metric.key" class="metric">
          <div class="metric-value">
            {{ formatNumber(metric.value) }}<small>{{ metric.unit }}</small>
          </div>
          <div class="bar-area">
            <div
              class="bar"
              :class="metric.color"
              :style="{
                height: `${Math.max(8, (metric.value / Math.max(totals[metric.key], 1)) * 100)}%`,
              }"
            ></div>
          </div>
          <strong>{{ metric.label }}</strong>
        </div>
      </div>
    </section>
    <footer>
      <span>DATENQUELLE · EIGENE FAHRTEN-DOKUMENTATION</span
      ><span>ROUTEN ALS LUFTLINIE · KARTENBASIS © OPENSTREETMAP</span>
    </footer>
  </main>
</template>

<style>
@import url("https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Space+Grotesk:wght@400;500;600;700&display=swap");
:root {
  font-family: "Space Grotesk", sans-serif;
  color: #16324f;
  background: #fff;
  font-synthesis: none;
  --ukraine-blue: #0057b7;
  --ukraine-blue-dark: #003b7a;
  --ukraine-yellow: #f2c230;
  --ink: #16324f;
  --muted: #5f6c74;
  --line: #d9dfe7;
}
* {
  box-sizing: border-box;
}
body {
  margin: 0;
  min-width: 320px;
}
.app-shell {
  max-width: 1440px;
  margin: auto;
  padding: 0 42px;
}
.topbar {
  height: 82px;
  border-bottom: 1px solid var(--line);
  display: flex;
  align-items: center;
  justify-content: space-between;
}
.brand {
  display: flex;
  align-items: center;
  gap: 10px;
  font-size: 17px;
  font-weight: 700;
  letter-spacing: -0.3px;
}
.brand-mark {
  display: grid;
  place-items: center;
  width: 30px;
  height: 40px;
  overflow: hidden;
  background: #fff;
}
.brand-mark img {
  display: block;
  width: 100%;
  height: 100%;
  object-fit: contain;
}
.brand small,
.topbar-meta,
.eyebrow,
footer {
  font-family: "DM Mono", monospace;
  font-size: 10px;
  letter-spacing: 1.6px;
}
.brand small {
  color: #5f6c74;
  margin-left: 12px;
}
.topbar-meta {
  display: flex;
  gap: 13px;
  align-items: center;
  color: #5f6c74;
}
.topbar-meta strong {
  color: #16324f;
  font-weight: 500;
}
.live-dot {
  width: 7px;
  height: 7px;
  border-radius: 50%;
  background: var(--ukraine-blue);
  box-shadow: 0 0 0 4px #dbe9f8;
}
.intro {
  display: flex;
  justify-content: space-between;
  align-items: end;
  padding: 66px 0 46px;
}
.eyebrow {
  color: var(--ukraine-blue);
  margin: 0 0 14px;
}
.intro h1 {
  font-size: clamp(42px, 5.5vw, 78px);
  line-height: 0.94;
  letter-spacing: -3px;
  margin: 0;
  font-weight: 600;
}
.intro h1 em {
  font-family: Georgia, serif;
  font-weight: 400;
  color: var(--ukraine-blue);
}
.intro-copy {
  max-width: 315px;
  color: #5f6c74;
  line-height: 1.55;
  margin: 0 5% 4px 0;
}
.dashboard {
  display: grid;
  grid-template-columns: minmax(0, 2.1fr) minmax(300px, 0.9fr);
  gap: 18px;
}
.panel,
.side-panel {
  background: #fff;
  border: 1px solid var(--line);
}
.map-panel {
  position: relative;
  min-height: 550px;
}
.map {
  height: 550px;
  z-index: 0;
}
.map-loading {
  position: absolute;
  inset: 0;
  display: grid;
  place-items: center;
  background: #eef5fbf2;
  z-index: 500;
  color: #16324f;
}
.map-overlay {
  position: absolute;
  bottom: 14px;
  left: 14px;
  right: 14px;
  z-index: 400;
  display: flex;
  gap: 18px;
  flex-wrap: wrap;
  padding: 10px 12px;
  background: #fffffffa;
  font:
    10px "DM Mono",
    monospace;
  color: #5f6c74;
}
.route-key {
  display: flex;
  align-items: center;
  gap: 6px;
}
.key-line {
  display: inline-block;
  width: 19px;
  border-top: 2px solid #4c6888;
}
.key-line.active {
  border-color: var(--ukraine-yellow);
}
.border-key {
  border-color: var(--ukraine-blue);
}
.front-key {
  border-color: #b11f36;
  border-top-style: dashed;
}
.layer-toggle {
  display: flex;
  align-items: center;
  gap: 6px;
  cursor: pointer;
  user-select: none;
}
.layer-toggle input {
  accent-color: #16324f;
  margin: 0;
}
.key-dot {
  width: 8px;
  height: 8px;
  border: 2px solid #16324f;
  border-radius: 50%;
  background: #f7f4ed;
}
.side-panel {
  padding: 31px 28px;
  display: flex;
  flex-direction: column;
  min-height: 550px;
}
.tour-heading {
  display: flex;
  justify-content: space-between;
  border-bottom: 1px solid var(--line);
  padding-bottom: 23px;
}
.tour-heading h2,
.section-heading h2 {
  font-size: 31px;
  letter-spacing: -1px;
  margin: 0;
}
.tour-count {
  font:
    12px "DM Mono",
    monospace;
  color: #8d8a80;
}
.date-range {
  font:
    12px "DM Mono",
    monospace;
  color: #5f6c74;
  margin: 26px 0;
}
.date-range span {
  color: var(--ukraine-blue);
  margin: 0 6px;
}
.route-list {
  display: flex;
  flex-wrap: wrap;
  gap: 0 15px;
  font-size: 18px;
  font-weight: 600;
  line-height: 1.8;
}
.route-list span:not(:last-child)::after {
  content: "→";
  color: var(--ukraine-blue);
  margin-left: 15px;
  font-weight: 400;
}
.route-list .destination {
  color: var(--ukraine-blue);
}
.note {
  font-size: 13px;
  line-height: 1.5;
  color: #6c777c;
  border-left: 3px solid var(--ukraine-yellow);
  padding-left: 12px;
  margin-top: 24px;
}
.front-status {
  display: flex;
  align-items: center;
  gap: 8px;
  color: #b11f36;
  font:
    10px "DM Mono",
    monospace;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}
.front-status span {
  width: 8px;
  height: 8px;
  border-radius: 50%;
  background: #b11f36;
}
.playback {
  margin-top: auto;
  display: grid;
  grid-template-columns: 39px 1fr;
  gap: 12px;
  align-items: center;
  border-top: 1px solid var(--line);
  padding-top: 22px;
}
.play-button {
  border: 0;
  background: var(--ukraine-yellow);
  color: var(--ukraine-blue-dark);
  width: 39px;
  height: 39px;
  font-size: 15px;
  cursor: pointer;
}
.progress-track {
  height: 3px;
  background: #e1dfd8;
}
.progress-track span {
  display: block;
  height: 100%;
  background: var(--ukraine-blue);
  transition: width 0.4s ease;
}
.play-label {
  grid-column: 2;
  font:
    10px "DM Mono",
    monospace;
  color: #8d8a80;
  text-transform: uppercase;
}
.tour-selector {
  display: grid;
  grid-template-columns: repeat(7, 1fr);
  gap: 5px;
  margin-top: 26px;
}
.tour-selector button {
  height: 30px;
  border: 1px solid #b8c7d8;
  background: transparent;
  color: #5f6c74;
  font:
    11px "DM Mono",
    monospace;
  cursor: pointer;
}
.tour-selector button.selected {
  background: var(--ukraine-blue);
  color: #fff;
  border-color: var(--ukraine-blue);
}
.metrics-section {
  padding: 72px 0 55px;
}
.section-heading {
  display: flex;
  justify-content: space-between;
  align-items: end;
  margin-bottom: 28px;
}
.section-heading > span {
  font:
    11px "DM Mono",
    monospace;
  color: #8d8a80;
}
.chart {
  height: 225px;
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 18px;
  border-bottom: 1px solid #16324f;
}
.metric {
  display: grid;
  grid-template-rows: 36px 1fr 29px;
  min-width: 0;
}
.metric-value {
  font:
    19px "DM Mono",
    monospace;
  color: #16324f;
}
.metric-value small {
  font-size: 11px;
  color: #8d8a80;
}
.bar-area {
  height: 160px;
  display: flex;
  align-items: end;
  background: linear-gradient(to top, #e5e2d9 1px, transparent 1px);
  background-size: 100% 40px;
}
.bar {
  width: 100%;
  min-height: 8px;
  transition: height 0.7s cubic-bezier(0.2, 0.8, 0.2, 1);
}
.bar.coral {
  background: var(--ukraine-blue);
}
.bar.blue {
  background: #4f86c6;
}
.bar.gold {
  background: var(--ukraine-yellow);
}
.bar.mint {
  background: #75a99c;
}
.bar.ink {
  background: #16324f;
}
.metric strong {
  font-size: 13px;
  font-weight: 500;
}
.status {
  padding: 60px 0;
}
.error {
  color: #b43727;
}
footer {
  display: flex;
  justify-content: space-between;
  border-top: 1px solid var(--line);
  padding: 21px 0 28px;
  color: #8d8a80;
  font-size: 9px;
  letter-spacing: 1px;
}
@media (max-width: 900px) {
  .app-shell {
    padding: 0 20px;
  }
  .intro {
    display: block;
    padding: 45px 0 32px;
  }
  .intro-copy {
    margin: 22px 0 0;
  }
  .dashboard {
    grid-template-columns: 1fr;
  }
  .map-panel,
  .map {
    min-height: 420px;
    height: 420px;
  }
  .side-panel {
    min-height: 470px;
  }
  .metrics-section {
    padding-top: 52px;
  }
  .chart {
    gap: 8px;
  }
  .metric-value {
    font-size: 14px;
  }
  .metric strong {
    font-size: 11px;
  }
  footer {
    display: block;
    line-height: 2;
  }
  .topbar-meta {
    font-size: 9px;
  }
  .brand small {
    display: none;
  }
}
@media (max-width: 520px) {
  .topbar-meta strong {
    display: none;
  }
  .intro h1 {
    letter-spacing: -2px;
  }
  .map-overlay {
    gap: 9px;
  }
  .route-key {
    font-size: 8px;
  }
  .section-heading {
    display: block;
  }
  .section-heading > span {
    display: block;
    margin-top: 10px;
  }
  .chart {
    gap: 5px;
  }
  .metric-value {
    font-size: 12px;
  }
  .metric-value small {
    font-size: 8px;
  }
  .metric strong {
    font-size: 10px;
  }
  .tour-selector {
    grid-template-columns: repeat(7, 1fr);
  }
}
</style>
