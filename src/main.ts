import './style.css';
import * as THREE from 'three';
import { OrbitControls } from 'three/examples/jsm/controls/OrbitControls.js';
import { CSG } from 'three-csg-ts';
import { inject } from '@vercel/analytics';

inject();

// ─── OLD FORMAT (walls[]) ────────────────────────────────────────────────────
interface WallData {
  class: string;
  confidence: number;
  bbox: { x1: number; y1: number; x2: number; y2: number };
  center: { x: number; y: number };
  area: number;
}
interface OldJsonData {
  walls: WallData[];
}

// ─── NEW FORMAT (Paradigm Estimate) ─────────────────────────────────────────
interface NewRecord {
  materialType: string;
  settings: {
    name?: string;
    type?: string;
    pitch?: string;
    area?: number;
    linear_total?: number;
    board_size?: string;
    oc_spacing?: string;
    direction?: string;
    height?: string;
    floor_level?: string;
    [key: string]: unknown;
  };
  coordinates_real_world: [number, number][];
  scale_factor_float: number;
}
interface NewJsonData {
  project_id: string;
  records: NewRecord[];
}

// ─── Wall registry entry ─────────────────────────────────────────────────────
interface WallEntry {
  pts: [number, number][];
  cx: number;
  cy: number;
  height: number;
  thickness: number;
  length: number;         // current world-space length of the wall (feet)
  baseElev: number;
  label: string;          // settings.name or fallback
  wallType: string;       // categorised type label
  record: NewRecord;
  originalColor: number;
  // World-space anchor — used by Apply so position never drifts
  worldPos: THREE.Vector3;
  worldRotY: number;
}

interface EstimateRow {
  id: number;
  wallLabel: string;
  wallType: string;
  volume: number;
  price: number;
}

// ─── SCALE helpers ───────────────────────────────────────────────────────────
const SCALE = 1 / 12; // inches → feet

const MATERIAL_COLORS: Record<string, number> = {
  roof_system: 0xf97316,
  eave_length: 0x22d3ee,
  ridge_length: 0xa855f7,
  hip_length: 0xfbbf24,
  valley_length: 0x60a5fa,
  gable_length: 0x4ade80,
  wall: 0x94a3b8,
  default: 0xffffff,
};

const LABEL_MAP: Record<string, string> = {
  roof_system: 'Roof Panel / Rafter',
  eave_length: 'Eave',
  ridge_length: 'Ridge',
  hip_length: 'Hip',
  valley_length: 'Valley',
  gable_length: 'Gable',
  wall: 'Wall',
};

// ─── Wall-type colour palette ────────────────────────────────────────────────
const WALL_TYPE_COLORS: Record<string, string> = {
  'Perimeter Wall': '#ef4444',
  'Interior Wall': '#3b82f6',
  'Foundation Wall': '#64748b',
  'Knee Wall': '#f59e0b',
  'Exterior Wall': '#10b981',
  'Other Wall': '#94a3b8',
};

// ─── ELEVATION helpers ───────────────────────────────────────────────────────
const ELEVATIONS: Record<string, number> = {
  roof_system: 0,
  eave_length: 0.05,
  valley_length: -0.05,
  hip_length: 0.15,
  ridge_length: 0.3,
  gable_length: 0.1,
  default: 0,
};

const FLOOR_ELEVATIONS: Record<string, number> = {
  'GARAGE': -1.0,
  'MAIN FLOOR': 0.0,
  'PORCH WALL': -0.5,
  'SECOND FLOOR': 9.0,
  'default': 0.0,
};

// ─── Categorise a wall's settings.type into a display label ──────────────────
function categoriseWallType(settings: NewRecord['settings']): string {
  const t = (settings.type || settings.name || '').toUpperCase();
  if (t.includes('PERIMETER')) return 'Perimeter Wall';
  if (t.includes('INTERIOR')) return 'Interior Wall';
  if (t.includes('FOUNDATION')) return 'Foundation Wall';
  if (t.includes('KNEE')) return 'Knee Wall';
  if (t.includes('EXTERIOR')) return 'Exterior Wall';
  return 'Other Wall';
}

// ─── Old-format class → display label ────────────────────────────────────────
const OLD_CLASS_LABEL: Record<string, string> = {
  perimeter_wall: 'Perimeter Wall',
  interior_wall: 'Interior Wall',
  foundation_wall: 'Foundation Wall',
  block_foundation_wall: 'Foundation Wall',
  knee_wall: 'Knee Wall',
};

class HouseViewer {
  private scene: THREE.Scene;
  private camera: THREE.PerspectiveCamera;
  private renderer: THREE.WebGLRenderer;
  private controls: OrbitControls;
  private buildingGroup: THREE.Group;

  // ─── Wall selection ─────────────────────────────────────────────────────
  private raycaster = new THREE.Raycaster();
  private pointer = new THREE.Vector2();
  /** Maps every wall mesh → its data so we can edit / re-render it */
  private wallRegistry = new Map<THREE.Mesh, WallEntry>();
  private selectedWall: THREE.Mesh | null = null;
  private estimateRows: EstimateRow[] = [];

  // ─── Interactive Cutout Placement ───────────────────────────────────────
  private placementMode: 'none' | 'door' | 'window' = 'none';
  private activeCutterMesh: THREE.Mesh | null = null;
  private activeDisplayMesh: THREE.Mesh | null = null;
  private activeWallTarget: THREE.Mesh | null = null;
  private cutterWidth = 0;
  private cutterHeight = 0;
  private cutterThickness = 2.0;
  private cutterSill = 0;

  constructor() {
    this.scene = new THREE.Scene();
    // White background as requested
    this.scene.background = new THREE.Color(0xffffff);

    const canvas = document.querySelector('#three-canvas') as HTMLCanvasElement;
    this.renderer = new THREE.WebGLRenderer({ canvas, antialias: true });
    this.renderer.setSize(window.innerWidth, window.innerHeight);
    this.renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    this.renderer.shadowMap.enabled = true;

    this.camera = new THREE.PerspectiveCamera(60, window.innerWidth / window.innerHeight, 0.1, 5000);
    this.camera.position.set(80, 80, 80);

    this.controls = new OrbitControls(this.camera, this.renderer.domElement);
    this.controls.enableDamping = true;

    this.buildingGroup = new THREE.Group();
    this.scene.add(this.buildingGroup);

    this.initLights();
    this.initGrid();
    this.initEventListeners();
    this.initModalListeners();
    this.initEstimateListeners();
    this.animate();
  }

  // ─── Lights ──────────────────────────────────────────────────────────────
  private initLights() {
    this.scene.add(new THREE.AmbientLight(0xffffff, 0.5));
    const sun = new THREE.DirectionalLight(0xfff8e7, 1.2);
    sun.position.set(100, 150, 80);
    sun.castShadow = true;
    this.scene.add(sun);
    const fill = new THREE.HemisphereLight(0x818cf8, 0x1e293b, 0.4);
    this.scene.add(fill);
  }

  private initGrid() {
    const grid = new THREE.GridHelper(500, 100, 0x1e293b, 0x1e293b);
    this.scene.add(grid);
  }

  // ─── Event listeners ─────────────────────────────────────────────────────
  private initEventListeners() {
    window.addEventListener('resize', () => {
      this.camera.aspect = window.innerWidth / window.innerHeight;
      this.camera.updateProjectionMatrix();
      this.renderer.setSize(window.innerWidth, window.innerHeight);
    });

    const fileInput = document.getElementById('json-upload') as HTMLInputElement;
    fileInput.addEventListener('change', (e) => this.handleFileUpload(e));

    // Click on canvas → try to select a wall
    const canvas = document.querySelector('#three-canvas') as HTMLCanvasElement;
    canvas.addEventListener('pointerdown', (e) => this.onCanvasClick(e));

    // Hover → show pointer cursor when over a wall
    canvas.addEventListener('pointermove', (e) => this.onCanvasHover(e));

    // Listen to cutout mode radio buttons
    document.querySelectorAll('input[name="cutout-mode"]').forEach(radio => {
      radio.addEventListener('change', (e) => {
        const val = (e.target as HTMLInputElement).value;
        this.placementMode = val as any;
        if (this.placementMode === 'none') {
          this.cancelCutoutPlacement();
        }
      });
    });

    // Listen to cutout toolbar
    const posXInput = document.getElementById('cutout-pos-x') as HTMLInputElement;
    const posYInput = document.getElementById('cutout-pos-y') as HTMLInputElement;
    posXInput.addEventListener('input', () => this.updateCutoutPosition());
    posYInput.addEventListener('input', () => this.updateCutoutPosition());

    document.getElementById('cutout-btn-cancel')?.addEventListener('click', () => {
      const resetRadio = document.getElementById('mode-none') as HTMLInputElement;
      if (resetRadio) resetRadio.checked = true;
      this.placementMode = 'none';
      this.cancelCutoutPlacement();
    });
    document.getElementById('cutout-btn-apply')?.addEventListener('click', () => this.applyCutout());
  }

  private initModalListeners() {
    const modal = document.getElementById('wall-modal')!;
    const card = document.getElementById('wall-modal-card')!;
    const btnClose = document.getElementById('wall-modal-close')!;
    const btnApply = document.getElementById('wall-modal-apply')!;
    const colorIn = document.getElementById('wall-color-input') as HTMLInputElement;

    // Close on backdrop click (outside the card)
    modal.addEventListener('pointerdown', (e) => {
      if (e.target === modal) this.closeModal();
    });
    btnClose.addEventListener('click', () => this.closeModal());
    btnApply.addEventListener('click', () => this.applyWallEdit());

    // Prevent card clicks from closing modal (prevent event bubbling)
    card.addEventListener('pointerdown', (e) => e.stopPropagation());

    colorIn.addEventListener('input', () => this.setModalColorControls(colorIn.value));
  }

  private initEstimateListeners() {
    const openBtn = document.getElementById('show-estimate-btn') as HTMLButtonElement;
    const modal = document.getElementById('estimate-modal') as HTMLElement;
    const card = document.getElementById('estimate-modal-card') as HTMLElement;
    const closeBtn = document.getElementById('estimate-close-btn') as HTMLButtonElement;
    const exportBtn = document.getElementById('estimate-export-btn') as HTMLButtonElement;

    openBtn.addEventListener('click', () => this.openEstimateModal());
    closeBtn.addEventListener('click', () => this.closeEstimateModal());
    exportBtn.addEventListener('click', () => this.exportEstimatePdf());

    modal.addEventListener('pointerdown', (e) => {
      if (e.target === modal) this.closeEstimateModal();
    });
    card.addEventListener('pointerdown', (e) => e.stopPropagation());

    this.updateEstimateButtonState();
  }

  private buildEstimateRows(): EstimateRow[] {
    let i = 1;
    return Array.from(this.wallRegistry.values()).map((entry) => {
      const volume = entry.length * entry.height * entry.thickness;
      return {
        id: i++,
        wallLabel: entry.label || `Wall ${i - 1}`,
        wallType: entry.wallType || 'Wall',
        volume,
        price: volume, // default: price equals volume in USD as requested
      };
    });
  }

  private formatUsd(value: number): string {
    return `$${value.toFixed(2)}`;
  }

  private escapeHtml(value: string): string {
    return value
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#39;');
  }

  private updateEstimateButtonState() {
    const openBtn = document.getElementById('show-estimate-btn') as HTMLButtonElement;
    const hasWalls = this.wallRegistry.size > 0;
    openBtn.disabled = !hasWalls;
    openBtn.style.opacity = hasWalls ? '1' : '0.5';
    openBtn.style.cursor = hasWalls ? 'pointer' : 'not-allowed';
  }

  private openEstimateModal() {
    this.estimateRows = this.buildEstimateRows();
    this.renderEstimateTable();
    const modal = document.getElementById('estimate-modal') as HTMLElement;
    modal.style.display = 'flex';
  }

  private closeEstimateModal() {
    const modal = document.getElementById('estimate-modal') as HTMLElement;
    modal.style.display = 'none';
  }

  private renderEstimateTable() {
    const tbody = document.getElementById('estimate-table-body') as HTMLElement;
    tbody.innerHTML = '';

    this.estimateRows.forEach((row) => {
      const tr = document.createElement('tr');
      const idTd = document.createElement('td');
      idTd.textContent = String(row.id);
      const labelTd = document.createElement('td');
      labelTd.textContent = row.wallLabel;
      const typeTd = document.createElement('td');
      typeTd.textContent = row.wallType;
      const volumeTd = document.createElement('td');
      volumeTd.textContent = row.volume.toFixed(2);
      const priceTd = document.createElement('td');
      const priceInput = document.createElement('input');
      priceInput.className = 'estimate-price-input';
      priceInput.type = 'number';
      priceInput.min = '0';
      priceInput.step = '0.01';
      priceInput.value = row.price.toFixed(2);
      priceInput.dataset.rowId = String(row.id);
      priceTd.appendChild(priceInput);

      tr.appendChild(idTd);
      tr.appendChild(labelTd);
      tr.appendChild(typeTd);
      tr.appendChild(volumeTd);
      tr.appendChild(priceTd);
      tbody.appendChild(tr);
    });

    tbody.querySelectorAll<HTMLInputElement>('.estimate-price-input').forEach((input) => {
      input.addEventListener('input', () => this.recalculateEstimateTotal());
    });

    this.recalculateEstimateTotal();
  }

  private recalculateEstimateTotal() {
    const totalEl = document.getElementById('estimate-final-total') as HTMLElement;
    let total = 0;

    document.querySelectorAll<HTMLInputElement>('.estimate-price-input').forEach((input) => {
      const parsed = parseFloat(input.value);
      if (!Number.isNaN(parsed) && parsed >= 0) total += parsed;
    });

    totalEl.textContent = this.formatUsd(total);
  }

  private refreshEstimateIfOpen() {
    const modal = document.getElementById('estimate-modal') as HTMLElement;
    if (modal.style.display !== 'none') {
      this.openEstimateModal();
    } else {
      this.updateEstimateButtonState();
    }
  }

  private exportEstimatePdf() {
    const rows: Array<{ id: string; label: string; type: string; volume: string; price: string }> = [];
    document.querySelectorAll('#estimate-table-body tr').forEach((tr) => {
      const tds = tr.querySelectorAll('td');
      if (tds.length < 5) return;
      const priceInput = tds[4].querySelector('input') as HTMLInputElement | null;
      rows.push({
        id: (tds[0].textContent || '').trim(),
        label: (tds[1].textContent || '').trim(),
        type: (tds[2].textContent || '').trim(),
        volume: (tds[3].textContent || '').trim(),
        price: ((priceInput?.value || '0')).trim(),
      });
    });

    const total = (document.getElementById('estimate-final-total')?.textContent || '$0.00').trim();
    const now = new Date().toLocaleString();

    const tableRows = rows.map((r) => `
      <tr>
        <td>${this.escapeHtml(r.id)}</td>
        <td>${this.escapeHtml(r.label)}</td>
        <td>${this.escapeHtml(r.type)}</td>
        <td>${this.escapeHtml(r.volume)}</td>
        <td>$${Number(r.price || '0').toFixed(2)}</td>
      </tr>
    `).join('');

    const html = `
      <html>
      <head>
        <title>Wall Estimate</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 24px; color: #111827; }
          h1 { margin: 0 0 10px; font-size: 22px; }
          p { margin: 0 0 16px; color: #4b5563; }
          table { width: 100%; border-collapse: collapse; margin-top: 12px; }
          th, td { border: 1px solid #d1d5db; padding: 8px; text-align: left; font-size: 13px; }
          th { background: #f3f4f6; }
          .total { margin-top: 16px; font-size: 18px; font-weight: 700; }
        </style>
      </head>
      <body>
        <h1>Bill Of Materials - Walls Estimate</h1>
        <p>Generated: ${now}</p>
        <table>
          <thead>
            <tr><th>#</th><th>Wall</th><th>Type</th><th>Volume (ft³)</th><th>Price ($)</th></tr>
          </thead>
          <tbody>${tableRows}</tbody>
        </table>
        <div class="total">Final Price: ${total}</div>
      </body>
      </html>
    `;

    const printWindow = window.open('', '_blank');
    if (!printWindow) return;
    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
  }

  private normalizeHexColor(rawValue: string): string | null {
    const trimmed = rawValue.trim();
    if (!trimmed) return null;

    const withHash = trimmed.startsWith('#') ? trimmed : `#${trimmed}`;
    const hexBody = withHash.slice(1);
    if (!/^[0-9a-fA-F]+$/.test(hexBody)) return null;

    if (hexBody.length === 3) {
      const expanded = hexBody.split('').map((c) => `${c}${c}`).join('');
      return `#${expanded.toLowerCase()}`;
    }
    if (hexBody.length === 6) {
      return `#${hexBody.toLowerCase()}`;
    }
    return null;
  }

  private setModalColorControls(hexValue: string) {
    const normalized = this.normalizeHexColor(hexValue);
    if (!normalized) return;

    const colorIn = document.getElementById('wall-color-input') as HTMLInputElement;
    colorIn.value = normalized;
  }

  // ─── Cutout Placement Logic ───────────────────────────────────────────────
  private startCutoutPlacement(targetMesh: THREE.Mesh) {
    if (this.placementMode === 'none') return;

    this.activeWallTarget = targetMesh;
    const entry = this.wallRegistry.get(targetMesh)!;

    this.cutterWidth = this.placementMode === 'window' ? 4 : 3;
    this.cutterHeight = this.placementMode === 'window' ? 3 : 7;
    this.cutterSill = this.placementMode === 'window' ? 3 : 0;
    this.cutterThickness = 2.0;

    // Create a visual indicator to show where it's being placed
    const geom = new THREE.BoxGeometry(this.cutterWidth, this.cutterHeight, this.cutterThickness);
    this.activeCutterMesh = new THREE.Mesh(geom, new THREE.MeshBasicMaterial({ color: 0xff0000, wireframe: true, transparent: true, opacity: 0.3 }));

    // Create the dummy display mesh
    let displayMat;
    if (this.placementMode === 'window') {
      displayMat = new THREE.MeshStandardMaterial({ color: 0x93c5fd, metalness: 0.18, roughness: 0.08, transparent: true, opacity: 0.5, side: THREE.DoubleSide });
    } else {
      displayMat = new THREE.MeshStandardMaterial({ color: 0x7c3aed, metalness: 0.05, roughness: 0.6, transparent: true, opacity: 0.5, side: THREE.DoubleSide });
    }
    this.activeDisplayMesh = new THREE.Mesh(new THREE.BoxGeometry(this.cutterWidth, this.cutterHeight, 0.45), displayMat);

    this.buildingGroup.add(this.activeCutterMesh);
    this.buildingGroup.add(this.activeDisplayMesh);

    // Setup the toolbar UI
    const toolbar = document.getElementById('cutout-toolbar')!;
    const title = document.getElementById('cutout-toolbar-title')!;
    const sliderX = document.getElementById('cutout-pos-x') as HTMLInputElement;
    const sliderY = document.getElementById('cutout-pos-y') as HTMLInputElement;
    const groupY = document.getElementById('cutout-pos-y-group')!;

    title.textContent = this.placementMode === 'window' ? 'Window' : 'Door';

    // Configure slider X (along wall length). Midpoint is 0. Maximum bounds are half length minus half cutter.
    const maxSlide = Math.max(0, (entry.length / 2) - (this.cutterWidth / 2));
    sliderX.min = (-maxSlide).toString();
    sliderX.max = maxSlide.toString();
    sliderX.value = "0";

    if (this.placementMode === 'window') {
      const maxY = Math.max(0, entry.height - this.cutterHeight);
      sliderY.min = "0";
      sliderY.max = maxY.toString();
      sliderY.value = this.cutterSill.toString();
      groupY.style.display = 'block';
    } else {
      // Doors are locked to the floor
      groupY.style.display = 'none';
      sliderY.value = "0";
    }

    toolbar.style.display = 'block';
    this.updateCutoutPosition();
  }

  private updateCutoutPosition() {
    if (!this.activeWallTarget || !this.activeCutterMesh || !this.activeDisplayMesh) return;

    const entry = this.wallRegistry.get(this.activeWallTarget)!;
    const sliderX = document.getElementById('cutout-pos-x') as HTMLInputElement;
    const sliderY = document.getElementById('cutout-pos-y') as HTMLInputElement;

    const dx = parseFloat(sliderX.value);
    const sill = parseFloat(sliderY.value);

    // Sliding horizontally along the wall means moving in the direction of rotY.
    const cosR = Math.cos(entry.worldRotY);
    const sinR = Math.sin(-entry.worldRotY);

    const targetX = entry.worldPos.x + (dx * cosR);
    const targetZ = entry.worldPos.z + (dx * sinR);
    const targetY = entry.baseElev + sill + (this.cutterHeight / 2);

    this.activeCutterMesh.position.set(targetX, targetY, targetZ);
    this.activeCutterMesh.rotation.y = entry.worldRotY;

    this.activeDisplayMesh.position.copy(this.activeCutterMesh.position);
    this.activeDisplayMesh.rotation.y = this.activeCutterMesh.rotation.y;
  }

  private cancelCutoutPlacement() {
    if (this.activeCutterMesh) {
      this.buildingGroup.remove(this.activeCutterMesh);
      this.activeCutterMesh.geometry.dispose();
      this.activeCutterMesh = null;
    }
    if (this.activeDisplayMesh) {
      this.buildingGroup.remove(this.activeDisplayMesh);
      this.activeDisplayMesh.geometry.dispose();
      this.activeDisplayMesh = null;
    }
    this.activeWallTarget = null;
    document.getElementById('cutout-toolbar')!.style.display = 'none';
  }

  private applyCutout() {
    if (!this.activeWallTarget || !this.activeCutterMesh || !this.activeDisplayMesh) return;

    const originalPos = this.activeWallTarget.position.clone();
    const originalQuat = this.activeWallTarget.quaternion.clone();
    const originalScale = this.activeWallTarget.scale.clone();

    this.activeCutterMesh.updateMatrixWorld(true);
    this.activeWallTarget.updateMatrixWorld(true);

    // Perform CSG subtraction on the wall
    const cutWallMesh = CSG.subtract(this.activeWallTarget, this.activeCutterMesh);
    cutWallMesh.material = this.activeWallTarget.material;
    const newGeometry = cutWallMesh.geometry.clone();
    newGeometry.computeBoundingBox();
    newGeometry.computeBoundingSphere();
    newGeometry.computeVertexNormals();

    // Replace actual wall geometry
    this.activeWallTarget.geometry.dispose();
    this.activeWallTarget.geometry = newGeometry;
    this.activeWallTarget.position.copy(originalPos);
    this.activeWallTarget.quaternion.copy(originalQuat);
    this.activeWallTarget.scale.copy(originalScale);
    this.activeWallTarget.updateMatrixWorld(true);

    // Display Mesh shadow config
    if (this.placementMode === 'door') {
      this.activeDisplayMesh.castShadow = true;
    }

    // Keep display mesh in scene permanently. Nullify pointers so cancel ignores it.
    this.buildingGroup.remove(this.activeCutterMesh);
    this.activeCutterMesh.geometry.dispose();
    this.activeCutterMesh = null;
    this.activeDisplayMesh = null;
    this.activeWallTarget = null;

    document.getElementById('cutout-toolbar')!.style.display = 'none';

    // Uncheck radio automatically so user must click it again to add another
    const resetRadio = document.getElementById('mode-none') as HTMLInputElement;
    if (resetRadio) resetRadio.checked = true;
    this.placementMode = 'none';
    this.refreshEstimateIfOpen();
  }

  // ─── File upload ─────────────────────────────────────────────────────────
  private async handleFileUpload(event: Event) {
    const target = event.target as HTMLInputElement;
    const file = target.files?.[0];
    if (!file) return;

    const fileInfo = document.querySelector('.file-info') as HTMLElement;
    fileInfo.textContent = `Loading: ${file.name}…`;

    try {
      const text = await file.text();
      const data = JSON.parse(text);

      if (data.records && Array.isArray(data.records)) {
        this.renderNewFormat(data as NewJsonData, file.name);
      } else if (data.walls && Array.isArray(data.walls)) {
        if (data.walls.length > 0 && data.walls[0].start && data.walls[0].end) {
          this.renderLineFormat(data, file.name);
        } else {
          this.renderOldFormat(data as OldJsonData, file.name);
        }
      } else {
        fileInfo.textContent = 'Error: Unrecognised JSON format.';
      }
    } catch (e) {
      fileInfo.textContent = 'Error: Could not parse JSON file.';
      console.error(e);
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // NEW FORMAT renderer
  // ═══════════════════════════════════════════════════════════════════════════
  private renderNewFormat(data: NewJsonData, _fileName: string) {
    this.clearScene();

    const records = data.records;
    if (!records || records.length === 0) return;

    // ── 1. global bounds ─────────────────────────────────────────────────
    let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
    records.forEach(rec => {
      rec.coordinates_real_world.forEach(([x, y]) => {
        minX = Math.min(minX, x); maxX = Math.max(maxX, x);
        minY = Math.min(minY, y); maxY = Math.max(maxY, y);
      });
    });
    const cx = (minX + maxX) / 2;
    const cy = (minY + maxY) / 2;

    // ── 2. counts ────────────────────────────────────────────────────────
    const typeCounts: Record<string, number> = {};
    const wallTypeCounts: Record<string, number> = {};

    // ── 3. render ────────────────────────────────────────────────────────
    records.forEach(rec => {
      const pts = rec.coordinates_real_world;
      if (!pts || pts.length === 0) return;

      const mt = rec.materialType || 'default';
      typeCounts[mt] = (typeCounts[mt] || 0) + 1;

      const color = MATERIAL_COLORS[mt] ?? MATERIAL_COLORS.default;
      const elev = ELEVATIONS[mt] ?? ELEVATIONS.default;
      const pitch = parseFloat(rec.settings.pitch || '4');

      const toV3 = ([x, y]: [number, number], z = 0): THREE.Vector3 =>
        new THREE.Vector3((x - cx) * SCALE, z, (y - cy) * SCALE);

      if (pts.length >= 3) {
        this.renderPolygon(pts, cx, cy, pitch, color, elev);
      } else if (pts.length === 2) {
        if (mt === 'wall') {
          // classify and count
          const wt = categoriseWallType(rec.settings);
          wallTypeCounts[wt] = (wallTypeCounts[wt] || 0) + 1;
          this.renderWall(pts, rec.settings, cx, cy, color, rec);
        } else {
          this.renderLine(toV3(pts[0], elev), toV3(pts[1], elev), color, mt);
        }
      }
    });

    // ── 4. stats panel ───────────────────────────────────────────────────
    this.updateStatsPanel(typeCounts, wallTypeCounts);
    this.addAutoFloorFromWalls();

    const fileInfo = document.querySelector('.file-info') as HTMLElement;
    fileInfo.textContent = `Project ${data.project_id} — ${records.length} records loaded`;

    this.frameCamera();
    this.updateLegendForNewFormat();
  }

  // ─── Polygon (roof panels) ────────────────────────────────────────────────
  private renderPolygon(
    pts: [number, number][],
    cx: number,
    cy: number,
    _pitch: number,
    color: number,
    baseElev: number
  ) {
    const shape = new THREE.Shape();
    const first = pts[0];
    shape.moveTo((first[0] - cx) * SCALE, (first[1] - cy) * SCALE);
    for (let i = 1; i < pts.length; i++) {
      shape.lineTo((pts[i][0] - cx) * SCALE, (pts[i][1] - cy) * SCALE);
    }
    shape.closePath();

    const geo = new THREE.ShapeGeometry(shape);
    const mat = new THREE.MeshStandardMaterial({
      color, metalness: 0.05, roughness: 0.7,
      side: THREE.DoubleSide, transparent: true, opacity: 0.82,
    });
    const mesh = new THREE.Mesh(geo, mat);
    mesh.rotation.x = -Math.PI / 2;
    mesh.position.y = baseElev;
    this.buildingGroup.add(mesh);

    const edgesGeo = new THREE.EdgesGeometry(geo);
    const edgesMat = new THREE.LineBasicMaterial({ color: 0xffffff, opacity: 0.3, transparent: true });
    const edges = new THREE.LineSegments(edgesGeo, edgesMat);
    edges.rotation.x = -Math.PI / 2;
    edges.position.y = baseElev + 0.01;
    this.buildingGroup.add(edges);
  }

  // ─── Line (eave / ridge / hip …) ─────────────────────────────────────────
  private renderLine(a: THREE.Vector3, b: THREE.Vector3, color: number, type: string) {
    const dir = new THREE.Vector3().subVectors(b, a);
    const length = dir.length();
    if (length < 0.001) return;

    const thickness = type === 'ridge_length' ? 0.25 : 0.15;
    const height = type === 'ridge_length' ? 0.3 : 0.12;

    const geo = new THREE.BoxGeometry(length, height, thickness);
    const mat = new THREE.MeshStandardMaterial({ color, metalness: 0.2, roughness: 0.6 });
    const mesh = new THREE.Mesh(geo, mat);

    mesh.position.copy(a).lerp(b, 0.5);
    mesh.position.y += height / 2;
    mesh.rotation.y = -Math.atan2(dir.z, dir.x);
    this.buildingGroup.add(mesh);
  }

  // ─── Wall ─────────────────────────────────────────────────────────────────
  private renderWall(
    pts: [number, number][],
    settings: NewRecord['settings'],
    cx: number,
    cy: number,
    color: number,
    record: NewRecord,
    existingMesh?: THREE.Mesh    // pass to rebuild in-place
  ): THREE.Mesh | null {
    const a = new THREE.Vector3((pts[0][0] - cx) * SCALE, 0, (pts[0][1] - cy) * SCALE);
    const b = new THREE.Vector3((pts[1][0] - cx) * SCALE, 0, (pts[1][1] - cy) * SCALE);

    const dir = new THREE.Vector3().subVectors(b, a);
    const length = dir.length();
    if (length < 0.001) return null;

    const floorLevel = (settings.floor_level as string) || 'default';
    const baseElev = FLOOR_ELEVATIONS[floorLevel] ?? FLOOR_ELEVATIONS.default;
    const height = parseFloat((settings.height as string) || '8');
    const thickness = 0.33; // ~4 inches in feet

    const geo = new THREE.BoxGeometry(length, height, thickness);
    const mat = new THREE.MeshStandardMaterial({
      color,
      metalness: 0.1, roughness: 0.8,
      emissive: new THREE.Color(0x000000),
    });

    const midPos = new THREE.Vector3().copy(a).lerp(b, 0.5);
    midPos.y = baseElev + height / 2;
    const rotY = -Math.atan2(dir.z, dir.x);

    let mesh: THREE.Mesh;
    if (existingMesh) {
      existingMesh.geometry.dispose();
      existingMesh.geometry = geo;
      existingMesh.position.copy(midPos);
      existingMesh.rotation.y = rotY;
      mesh = existingMesh;
    } else {
      mesh = new THREE.Mesh(geo, mat);
      mesh.position.copy(midPos);
      mesh.rotation.y = rotY;
      this.buildingGroup.add(mesh);
    }

    // Register / update registry
    const wallType = categoriseWallType(settings);
    this.wallRegistry.set(mesh, {
      pts, cx, cy, height, thickness, length,
      baseElev, label: settings.name || 'Wall', wallType, record,
      originalColor: color,
      worldPos: midPos.clone(),
      worldRotY: rotY,
    });

    return mesh;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // LINE FORMAT renderer (e.g. 2bhk.json)
  // ═══════════════════════════════════════════════════════════════════════════
  private renderLineFormat(data: any, _fileName: string) {
    this.clearScene();

    const walls = data.walls || [];
    if (!walls.length) return;

    let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;

    walls.forEach((w: any) => {
      if (w.start && w.end) {
        minX = Math.min(minX, w.start.x, w.end.x);
        maxX = Math.max(maxX, w.start.x, w.end.x);
        minY = Math.min(minY, w.start.y, w.end.y);
        maxY = Math.max(maxY, w.start.y, w.end.y);
      }
    });

    if (!isFinite(minX)) return;
    const cx = (minX + maxX) / 2;
    const cy = (minY + maxY) / 2;

    const wallTypeCounts: Record<string, number> = {};
    const wallColor = 0x94a3b8;

    let scaleFactor = 1;
    if (data.units === 'inches') {
      scaleFactor = SCALE; // 1/12
    } else if (data.units === 'feet') {
      scaleFactor = 1;
    } else if (data.units === 'meters' || data.units === 'metres') {
      scaleFactor = 3.28084;
    }

    const doorColor = 0x7c3aed;
    const windowGlass = 0x93c5fd;

    // 1. Prepare cutout meshes
    const cutouts = data.cutouts || [];
    const cutoutDataList: { mesh: THREE.Mesh, isWindow: boolean }[] = [];

    cutouts.forEach((cutout: any) => {
      if (!cutout.position) return;

      const x = cutout.position.x;
      const y = cutout.position.y;

      const cutoutWidth = (cutout.width || 3) * scaleFactor;
      const cutoutHeight = (cutout.height || 7) * scaleFactor;
      // Make cutout thickness much larger than typical wall to guarantee a clean cut completely through
      const thickness = 2.0;

      let rotY = 0;
      // Attempt to orient the cutout along the wall it sits on
      walls.forEach((w: any) => {
        if (!w.start || !w.end) return;
        // Simple distance-to-point check
        const distToStart = Math.hypot(x - w.start.x, y - w.start.y);
        const distToEnd = Math.hypot(x - w.end.x, y - w.end.y);
        const wallLen = Math.hypot(w.end.x - w.start.x, w.end.y - w.start.y);
        // If x,y is roughly on the segment
        if (distToStart + distToEnd <= wallLen + 0.1) {
          rotY = -Math.atan2(w.end.y - w.start.y, w.end.x - w.start.x);
        }
      });

      const isWindow = cutout.type === 'window';
      const sillHeight = isWindow ? (cutout.sill_height || 3) * scaleFactor : 0;

      const geom = new THREE.BoxGeometry(cutoutWidth, cutoutHeight, thickness);

      let mat;
      if (isWindow) {
        mat = new THREE.MeshStandardMaterial({ color: windowGlass, metalness: 0.18, roughness: 0.08, transparent: true, opacity: 0.5, side: THREE.DoubleSide });
      } else {
        mat = new THREE.MeshStandardMaterial({ color: doorColor, metalness: 0.05, roughness: 0.6 });
      }

      const mesh = new THREE.Mesh(geom, mat);
      const midPos = new THREE.Vector3((x - cx) * scaleFactor, sillHeight + cutoutHeight / 2, (y - cy) * scaleFactor);
      mesh.position.copy(midPos);
      mesh.rotation.y = rotY;

      mesh.castShadow = !isWindow;
      mesh.receiveShadow = true;

      mesh.updateMatrix(); // bake matrix for CSG
      cutoutDataList.push({ mesh, isWindow });

      // We still add the window glass to the scene!
      // But we will use the same geometry/mesh to subtract from wall. 
      // Actually we should clone the mesh for the visible window so we can make the cutter bigger.
      if (isWindow) {
        const displayMesh = mesh.clone();
        // Thin it back out for display
        displayMesh.geometry = new THREE.BoxGeometry(cutoutWidth, cutoutHeight, 0.45);
        this.buildingGroup.add(displayMesh);
      }
    });

    // 2. Build walls and subtract cutouts
    walls.forEach((w: any) => {
      if (!w.start || !w.end) return;

      const x1 = w.start.x;
      const y1 = w.start.y;
      const x2 = w.end.x;
      const y2 = w.end.y;

      const dx = x2 - x1;
      const dy = y2 - y1;
      const length = Math.sqrt(dx * dx + dy * dy);

      const actualLength = length * scaleFactor;
      const wallHeight = (w.height || data.default_wall_height || 11) * scaleFactor;
      const thickness = 0.5; // approx 6 inches

      const geom = new THREE.BoxGeometry(actualLength, wallHeight, thickness);
      const mat = new THREE.MeshStandardMaterial({ color: wallColor, flatShading: true, metalness: 0.0, roughness: 0.6 });
      let wallMesh: THREE.Mesh = new THREE.Mesh(geom, mat);

      const centerXpx = (x1 + x2) / 2;
      const centerYpx = (y1 + y2) / 2;

      // Note: y and z were swapped in THREE.js relative to typical 2D maps
      const midPos = new THREE.Vector3((centerXpx - cx) * scaleFactor, wallHeight / 2, (centerYpx - cy) * scaleFactor);
      wallMesh.position.copy(midPos);

      // We negate the angle because z grows "forward" towards camera
      const rotY = -Math.atan2(dy, dx);
      wallMesh.rotation.y = rotY;
      wallMesh.updateMatrix();

      // CSG Subtraction
      /*
      cutoutDataList.forEach(cutoutData => {
        // Rough bounding box intersection optimization could go here.
        // For safety, we use CSG on all walls against all cutouts since the lists are small.
        const cutWallMesh = CSG.subtract(wallMesh, cutoutData.mesh);
        // keep material
        cutWallMesh.material = mat;

        // After subtract the matrix is baked in, so reset position/rotation
        cutWallMesh.position.set(0, 0, 0);
        cutWallMesh.rotation.set(0, 0, 0);
        wallMesh = cutWallMesh;
        wallMesh.castShadow = true;
        wallMesh.receiveShadow = true;
      });
      */

      wallMesh.castShadow = true;
      wallMesh.receiveShadow = true;

      this.buildingGroup.add(wallMesh);

      const wallTypeLabel = w.room || 'Wall';
      wallTypeCounts[wallTypeLabel] = (wallTypeCounts[wallTypeLabel] || 0) + 1;

      const syntheticPts: [number, number][] = [
        [(x1 - cx) * scaleFactor, (y1 - cy) * scaleFactor],
        [(x2 - cx) * scaleFactor, (y2 - cy) * scaleFactor],
      ];

      const fakeRecord: NewRecord = {
        materialType: 'wall',
        settings: { name: wallTypeLabel, type: w.room, height: String(wallHeight), floor_level: 'default' },
        coordinates_real_world: syntheticPts,
        scale_factor_float: scaleFactor,
      } as NewRecord;

      this.wallRegistry.set(wallMesh, {
        pts: syntheticPts,
        cx, cy,
        height: wallHeight,
        thickness,
        length: actualLength,
        baseElev: 0,
        label: wallTypeLabel,
        wallType: wallTypeLabel,
        record: fakeRecord,
        originalColor: wallColor,
        worldPos: midPos.clone(),
        worldRotY: rotY,
      });
    });

    this.updateStatsPanel({}, wallTypeCounts);
    this.addAutoFloorFromWalls();

    const fileInfo = document.querySelector('.file-info') as HTMLElement;
    fileInfo.textContent = `Applied: ${_fileName}`;

    this.frameCamera();
    this.updateLegendForOldFormat();
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // OLD FORMAT renderer
  // ═══════════════════════════════════════════════════════════════════════════
  private renderOldFormat(data: any, _fileName: string) {
    this.clearScene();

    const walls = data.walls || [];
    const floors = data.floors || [];
    const ceilings = data.ceilings || [];
    const doors = data.doors || [];
    const windows = data.windows || [];

    if (!walls.length && !floors.length && !ceilings.length && !doors.length && !windows.length) return;

    // Compute image-space bounds so we can center the building at world origin.
    let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
    const extendFromBbox = (bbox: any) => {
      if (!bbox) return;
      const x1 = Number(bbox.x1 ?? 0);
      const y1 = Number(bbox.y1 ?? 0);
      const x2 = Number(bbox.x2 ?? 0);
      const y2 = Number(bbox.y2 ?? 0);
      minX = Math.min(minX, x1, x2);
      maxX = Math.max(maxX, x1, x2);
      minY = Math.min(minY, y1, y2);
      maxY = Math.max(maxY, y1, y2);
    };

    walls.forEach((w: any) => extendFromBbox(w.bbox));
    floors.forEach((f: any) => extendFromBbox(f.bbox));
    ceilings.forEach((c: any) => extendFromBbox(c.bbox));
    doors.forEach((d: any) => extendFromBbox(d.bbox));
    windows.forEach((w: any) => extendFromBbox(w.bbox));

    // Guard: if bounds never set, bail
    if (!isFinite(minX)) return;
    const cx = (minX + maxX) / 2;
    const cy = (minY + maxY) / 2;

    // Count classes for stats panel
    const classCounts: Record<string, number> = {};
    const wallTypeCounts: Record<string, number> = {};

    // Render floors/ceilings — non-selectable background
    // Use a tall interior height: set ceiling at 20 ft so the space feels larger.
    const defaultFloorHeight = 0; // ground
    const defaultCeilingHeight = 20; // ft (user-requested larger ceiling)
    const floorColor = 0x0b1220; // dark subtle
    const ceilingColor = 0xe6eef8; // light subtle so ceiling is visible against background
    const floorThicknessFt = 2;   // floor thickness in feet
    const ceilingThicknessFt = 1; // ceiling thickness in feet

    floors.forEach((f: any) => {
      const bbox = f.bbox || {};
      const x1 = Number(bbox.x1 ?? 0);
      const y1 = Number(bbox.y1 ?? 0);
      const x2 = Number(bbox.x2 ?? 0);
      const y2 = Number(bbox.y2 ?? 0);
      const width = Math.max(1, Math.abs(x2 - x1)) * SCALE;
      const depth = Math.max(1, Math.abs(y2 - y1)) * SCALE;

      // Create a thin box for the floor so it has thickness
      const geo = new THREE.BoxGeometry(width, floorThicknessFt, depth);
      // Floor transparency: 30% transparent => opacity = 0.7
      const mat = new THREE.MeshStandardMaterial({ color: floorColor, side: THREE.DoubleSide, transparent: true, opacity: 0.7, roughness: 0.85, metalness: 0.02 });
      const mesh = new THREE.Mesh(geo, mat);
      // Box geometry is centered — place it so top of floor sits at defaultFloorHeight
      mesh.position.set(((x1 + x2) / 2 - cx) * SCALE, defaultFloorHeight + (floorThicknessFt / 2), ((y1 + y2) / 2 - cy) * SCALE);
      mesh.receiveShadow = true;
      mesh.castShadow = false;
      this.buildingGroup.add(mesh);
    });

    // Render ceilings (non-selectable)
    ceilings.forEach((c: any) => {
      const bbox = c.bbox || {};
      const x1 = Number(bbox.x1 ?? 0);
      const y1 = Number(bbox.y1 ?? 0);
      const x2 = Number(bbox.x2 ?? 0);
      const y2 = Number(bbox.y2 ?? 0);
      const width = Math.max(1, Math.abs(x2 - x1)) * SCALE;
      const depth = Math.max(1, Math.abs(y2 - y1)) * SCALE;

      // Create a box for the ceiling so it has thickness
      const geo = new THREE.BoxGeometry(width, ceilingThicknessFt, depth);
      // Ceiling transparency: 50% transparent => opacity = 0.5
      const mat = new THREE.MeshStandardMaterial({ color: ceilingColor, side: THREE.DoubleSide, transparent: true, opacity: 0.5, roughness: 0.9, metalness: 0.0 });
      const mesh = new THREE.Mesh(geo, mat);
      // Place the ceiling so its top aligns at defaultCeilingHeight (center lowered by half thickness)
      mesh.position.set(((x1 + x2) / 2 - cx) * SCALE, defaultCeilingHeight - (ceilingThicknessFt / 2), ((y1 + y2) / 2 - cy) * SCALE);
      mesh.receiveShadow = true;
      mesh.castShadow = false;
      this.buildingGroup.add(mesh);
    });

    // Helper to compute orientation angle from various sources
    const computeAngle = (obj: any, widthPx: number, heightPx: number) => {
      let angleRad = 0;
      try {
        if (Array.isArray(obj.centerline_image) && obj.centerline_image.length >= 2) {
          const a = obj.centerline_image[0];
          const b = obj.centerline_image[1];
          if (Array.isArray(a) && Array.isArray(b)) angleRad = Math.atan2(b[1] - a[1], b[0] - a[0]);
        } else if (Array.isArray(obj.polygon_image) && obj.polygon_image.length >= 2) {
          const a = obj.polygon_image[0];
          const b = obj.polygon_image[1];
          if (Array.isArray(a) && Array.isArray(b)) angleRad = Math.atan2(b[1] - a[1], b[0] - a[0]);
        } else if (typeof obj._angle_deg === 'number' && !Number.isNaN(obj._angle_deg)) {
          angleRad = (obj._angle_deg * Math.PI) / 180;
        } else {
          angleRad = widthPx >= heightPx ? 0 : Math.PI / 2;
        }
      } catch {
        angleRad = widthPx >= heightPx ? 0 : Math.PI / 2;
      }
      return angleRad;
    };

    // Render walls and register them for selection/editing
    walls.forEach((w: any) => {
      const bbox = w.bbox || {};
      const x1 = Number(bbox.x1 ?? 0);
      const y1 = Number(bbox.y1 ?? 0);
      const x2 = Number(bbox.x2 ?? 0);
      const y2 = Number(bbox.y2 ?? 0);
      const widthPx = Math.abs(x2 - x1) || 1;
      const heightPx = Math.abs(y2 - y1) || 1;

      const angleRad = computeAngle(w, widthPx, heightPx);

      // world sizes
      const lengthPx = Math.max(widthPx, heightPx) || (w._length_px ?? 1);
      const length = lengthPx * SCALE;
      // Default wall height: slightly shorter than ceiling so there's a small gap for molding
      const wallHeight = (w._height_ft ?? Math.max(defaultCeilingHeight - 0.5, 1));
      const thicknessPx = Number(w._thickness_px ?? 10);
      const thickness = Math.max(0.1, thicknessPx * SCALE);

      const geom = new THREE.BoxGeometry(length, wallHeight, thickness);
      const colorKey = (w.class && MATERIAL_COLORS && MATERIAL_COLORS[w.class]) ? w.class : 'perimeter_wall';
      const color = (MATERIAL_COLORS && MATERIAL_COLORS[colorKey]) ? MATERIAL_COLORS[colorKey] : 0x94a3b8;
      const mat = new THREE.MeshStandardMaterial({ color, flatShading: true, metalness: 0.0, roughness: 0.6 });
      const mesh = new THREE.Mesh(geom, mat);

      // Position relative to computed center so buildingGroup is centred at origin
      const centerXpx = (x1 + x2) / 2;
      const centerYpx = (y1 + y2) / 2;
      const midPos = new THREE.Vector3((centerXpx - cx) * SCALE, wallHeight / 2, (centerYpx - cy) * SCALE);
      mesh.position.copy(midPos);

      const rotY = -angleRad;
      mesh.rotation.y = rotY;

      mesh.castShadow = true;
      mesh.receiveShadow = true;

      this.buildingGroup.add(mesh);

      // Register wall for selection/editing similar to new-format registry
      const wallTypeLabel = OLD_CLASS_LABEL[w.class] || w.class || 'Wall';
      classCounts[w.class] = (classCounts[w.class] || 0) + 1;
      wallTypeCounts[wallTypeLabel] = (wallTypeCounts[wallTypeLabel] || 0) + 1;

      // synthetic world-space pts (two endpoints) for consistency with NewRecord shape
      const syntheticPts: [number, number][] = [
        [(x1 - cx) * SCALE, (y1 - cy) * SCALE],
        [(x2 - cx) * SCALE, (y2 - cy) * SCALE],
      ];

      const fakeRecord: NewRecord = {
        materialType: 'wall',
        settings: { name: wallTypeLabel, type: w.class, height: String(wallHeight), floor_level: 'default' },
        coordinates_real_world: syntheticPts,
        scale_factor_float: SCALE,
      } as NewRecord;

      this.wallRegistry.set(mesh, {
        pts: syntheticPts,
        cx, cy,
        height: wallHeight,
        thickness,
        length,
        baseElev: 0,
        label: wallTypeLabel,
        wallType: wallTypeLabel,
        record: fakeRecord,
        originalColor: color,
        worldPos: midPos.clone(),
        worldRotY: rotY,
      });
    });

    // Render doors (non-selectable)
    const doorColor = 0x7c3aed; // subtle brown/purple for door fill
    doors.forEach((d: any) => {
      const bbox = d.bbox || {};
      const x1 = Number(bbox.x1 ?? 0);
      const y1 = Number(bbox.y1 ?? 0);
      const x2 = Number(bbox.x2 ?? 0);
      const y2 = Number(bbox.y2 ?? 0);
      const widthPx = Math.abs(x2 - x1) || 1;
      const heightPx = Math.abs(y2 - y1) || 1;
      const angleRad = computeAngle(d, widthPx, heightPx);

      const length = Math.max(widthPx, heightPx) * SCALE;
      const doorHeight = d.height_ft ?? 7; // default 7 ft
      const thickness = 0.2;

      const geo = new THREE.BoxGeometry(length, doorHeight, thickness);
      const mat = new THREE.MeshStandardMaterial({ color: doorColor, metalness: 0.05, roughness: 0.6 });
      const mesh = new THREE.Mesh(geo, mat);
      const centerXpx = (x1 + x2) / 2;
      const centerYpx = (y1 + y2) / 2;
      mesh.position.set((centerXpx - cx) * SCALE, doorHeight / 2, (centerYpx - cy) * SCALE);
      mesh.rotation.y = -angleRad;
      mesh.castShadow = true;
      mesh.receiveShadow = false;
      this.buildingGroup.add(mesh);
    });

    // Render windows (glass + frames)
    const windowGlass = 0x93c5fd; // light blue glass
    const frameColor = 0x374151; // dark gray frame
    windows.forEach((wi: any) => {
      const bbox = wi.bbox || {};
      const x1 = Number(bbox.x1 ?? 0);
      const y1 = Number(bbox.y1 ?? 0);
      const x2 = Number(bbox.x2 ?? 0);
      const y2 = Number(bbox.y2 ?? 0);
      const widthPx = Math.abs(x2 - x1) || 1;
      const heightPx = Math.abs(y2 - y1) || 1;
      const angleRad = computeAngle(wi, widthPx, heightPx);

      const length = Math.max(widthPx, heightPx) * SCALE;
      const winHeight = wi.height_ft ?? 4; // default window height
      const thickness = 0.15;
      // Glass pane
      const geo = new THREE.BoxGeometry(length, winHeight, thickness);
      // More glass-like: slightly reflective, low roughness, semi-transparent
      const mat = new THREE.MeshStandardMaterial({ color: windowGlass, metalness: 0.18, roughness: 0.08, transparent: true, opacity: 0.5, side: THREE.DoubleSide });
      const mesh = new THREE.Mesh(geo, mat);
      const centerXpx = (x1 + x2) / 2;
      const centerYpx = (y1 + y2) / 2;
      const sill = wi.sill_ft ?? 3; // default sill height
      mesh.position.set((centerXpx - cx) * SCALE, sill + winHeight / 2, (centerYpx - cy) * SCALE);
      mesh.rotation.y = -angleRad;
      mesh.castShadow = false;
      mesh.receiveShadow = true;
      this.buildingGroup.add(mesh);

      // Frame: four thin boxes around the glass pane
      const frameThickness = 0.08; // ft (about 1 inch)
      const frameDepth = thickness + 0.02;
      const halfW = length / 2;
      const halfH = winHeight / 2;

      // Helper to create a frame element at local (x, y)
      const makeFrame = (offsetX: number, offsetY: number, w: number, h: number) => {
        const fgeo = new THREE.BoxGeometry(w, h, frameDepth);
        const fmat = new THREE.MeshStandardMaterial({ color: frameColor, metalness: 0.05, roughness: 0.6 });
        const fmesh = new THREE.Mesh(fgeo, fmat);
        fmesh.position.set((centerXpx - cx) * SCALE + offsetX, sill + winHeight / 2 + offsetY, (centerYpx - cy) * SCALE);
        fmesh.rotation.y = -angleRad;
        fmesh.castShadow = true;
        fmesh.receiveShadow = true;
        this.buildingGroup.add(fmesh);
      };

      // Left vertical
      makeFrame(-halfW + frameThickness / 2, 0, frameThickness, winHeight + frameThickness);
      // Right vertical
      makeFrame(halfW - frameThickness / 2, 0, frameThickness, winHeight + frameThickness);
      // Top horizontal
      makeFrame(0, halfH - frameThickness / 2, length + frameThickness, frameThickness);
      // Bottom horizontal
      makeFrame(0, -halfH + frameThickness / 2, length + frameThickness, frameThickness);
    });

    // Update stats panel (old-format doesn't have material type counts)
    // Include floors/ceilings/doors/windows counts in the typeCounts passed to the UI so they're visible
    const extraTypeCounts: Record<string, number> = {};
    if (floors.length) extraTypeCounts['floor'] = floors.length;
    if (ceilings.length) extraTypeCounts['ceiling'] = ceilings.length;
    if (doors.length) extraTypeCounts['door'] = doors.length;
    if (windows.length) extraTypeCounts['window'] = windows.length;

    this.updateStatsPanel(extraTypeCounts, wallTypeCounts);
    if (!floors.length) this.addAutoFloorFromWalls();

    const fileInfo = document.querySelector('.file-info') as HTMLElement;
    fileInfo.textContent = `Applied: ${_fileName}`;

    this.frameCamera();
    this.updateLegendForOldFormat();
  }
  // private renderOldFormat(data: OldJsonData, _fileName: string) {
  //   this.clearScene();

  //   const walls = data.walls;
  //   if (!walls || walls.length === 0) return;

  //   let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
  //   walls.forEach(w => {
  //     minX = Math.min(minX, w.bbox.x1, w.bbox.x2);
  //     maxX = Math.max(maxX, w.bbox.x1, w.bbox.x2);
  //     minY = Math.min(minY, w.bbox.y1, w.bbox.y2);
  //     maxY = Math.max(maxY, w.bbox.y1, w.bbox.y2);
  //   });

  //   const cx = (minX + maxX) / 2;
  //   const cy = (minY + maxY) / 2;
  //   const size = Math.max(maxX - minX, maxY - minY);
  //   const sc = 50 / size;
  //   const wallH = 2.5;

  //   // Count by class
  //   const classCounts: Record<string, number> = {};

  //   walls.forEach(w => {
  //     classCounts[w.class] = (classCounts[w.class] || 0) + 1;

  //     const width = Math.abs(w.bbox.x2 - w.bbox.x1) * sc || 0.1;
  //     const depth = Math.abs(w.bbox.y2 - w.bbox.y1) * sc || 0.1;
  //     const baseColor = (this.wallMaterials[w.class] as THREE.MeshStandardMaterial)?.color?.getHex() ?? 0xffffff;
  //     const geo = new THREE.BoxGeometry(width, wallH, depth);
  //     const mat = new THREE.MeshStandardMaterial({
  //       color: baseColor, metalness: 0.1, roughness: 0.8,
  //       emissive: new THREE.Color(0x000000),
  //     });
  //     const mesh = new THREE.Mesh(geo, mat);
  //     const midPos = new THREE.Vector3((w.center.x - cx) * sc, wallH / 2, (w.center.y - cy) * sc);
  //     mesh.position.copy(midPos);
  //     // For old format, walls are axis-aligned, so rotation is 0 or PI/2
  //     const rotY = (width > depth) ? 0 : Math.PI / 2;
  //     mesh.rotation.y = rotY;
  //     this.buildingGroup.add(mesh);

  //     // Register for click-to-select
  //     const wallTypeLabel = OLD_CLASS_LABEL[w.class] || w.class;
  //     const syntheticPts: [number, number][] = [
  //       [w.bbox.x1, w.bbox.y1],
  //       [w.bbox.x2, w.bbox.y2],
  //     ];
  //     const fakeRecord: NewRecord = {
  //       materialType: 'wall',
  //       settings: { name: wallTypeLabel, type: w.class, height: String(wallH), floor_level: 'default' },
  //       coordinates_real_world: syntheticPts,
  //       scale_factor_float: sc,
  //     };
  //     this.wallRegistry.set(mesh, {
  //       pts: syntheticPts, cx, cy,
  //       height: wallH, thickness: depth, length: Math.max(width, depth), baseElev: 0,
  //       label: wallTypeLabel, wallType: wallTypeLabel,
  //       record: fakeRecord, originalColor: baseColor,
  //       worldPos: midPos.clone(),
  //       worldRotY: rotY,
  //     });
  //   });

  //   // Build wall type count for stats panel
  //   const wallTypeCounts: Record<string, number> = {};
  //   Object.entries(classCounts).forEach(([cls, count]) => {
  //     const label = OLD_CLASS_LABEL[cls] || cls;
  //     wallTypeCounts[label] = (wallTypeCounts[label] || 0) + count;
  //   });

  //   const statsPanel = document.getElementById('stats-panel') as HTMLElement;
  //   statsPanel.style.display = 'block';

  //   const wallSummary = document.getElementById('wall-summary') as HTMLElement;
  //   const wallSummaryList = document.getElementById('wall-summary-list') as HTMLElement;
  //   wallSummaryList.innerHTML = '';
  //   Object.entries(wallTypeCounts).forEach(([label, count]) => {
  //     const color = WALL_TYPE_COLORS[label] || '#94a3b8';
  //     wallSummaryList.innerHTML += `
  //       <div class="stat-row">
  //         <span class="stat-dot" style="background:${color}"></span>
  //         ${label}: <strong style="color:#f8fafc;margin-left:auto">${count}</strong>
  //       </div>`;
  //   });
  //   wallSummary.style.display = wallSummaryList.innerHTML ? 'block' : 'none';

  //   const statsList = document.getElementById('stats-list') as HTMLElement;
  //   statsList.innerHTML = `<div class="stat-row">Total Walls: <strong style="color:#f8fafc;margin-left:auto">${walls.length}</strong></div>`;

  //   const fileInfo = document.querySelector('.file-info') as HTMLElement;
  //   fileInfo.textContent = `Applied: ${_fileName}`;

  //   this.frameCamera();
  //   this.updateLegendForOldFormat();
  // }

  // ─── Stats Panel (new format) ─────────────────────────────────────────────
  private updateStatsPanel(
    typeCounts: Record<string, number>,
    wallTypeCounts: Record<string, number>
  ) {
    const statsPanel = document.getElementById('stats-panel') as HTMLElement;
    statsPanel.style.display = 'block';

    // Wall summary breakdown section
    const wallSummary = document.getElementById('wall-summary') as HTMLElement;
    const wallSummaryList = document.getElementById('wall-summary-list') as HTMLElement;
    wallSummaryList.innerHTML = '';

    const totalWalls = Object.values(wallTypeCounts).reduce((a, b) => a + b, 0);
    if (totalWalls > 0) {
      Object.entries(wallTypeCounts).forEach(([label, count]) => {
        const color = WALL_TYPE_COLORS[label] || '#94a3b8';
        const row = document.createElement('div');
        row.className = 'stat-row';
        row.innerHTML = `
          <span class="stat-dot" style="background:${color}"></span>
          ${label}
          <strong style="color:#f8fafc;margin-left:auto">${count}</strong>`;
        wallSummaryList.appendChild(row);
      });
      wallSummary.style.display = 'block';
    } else {
      wallSummary.style.display = 'none';
    }

    // General record breakdown
    const statsList = document.getElementById('stats-list') as HTMLElement;
    statsList.innerHTML = '';
    Object.entries(typeCounts).forEach(([type, count]) => {
      const div = document.createElement('div');
      div.className = 'stat-row';
      const dot = document.createElement('span');
      dot.className = 'stat-dot';
      dot.style.background = '#' + (MATERIAL_COLORS[type] ?? MATERIAL_COLORS.default).toString(16).padStart(6, '0');
      div.appendChild(dot);
      const label = LABEL_MAP[type] ?? type;
      div.innerHTML += `${label}<strong style="color:#f8fafc;margin-left:auto">${count}</strong>`;
      statsList.appendChild(div);
    });

    this.updateEstimateButtonState();
  }

  // ─── Raycasting ───────────────────────────────────────────────────────────
  private getWallMeshes(): THREE.Mesh[] {
    return Array.from(this.wallRegistry.keys());
  }

  private onCanvasClick(e: PointerEvent) {
    if (e.button !== 0) return; // left click only

    const canvas = this.renderer.domElement;
    const rect = canvas.getBoundingClientRect();
    this.pointer.x = ((e.clientX - rect.left) / rect.width) * 2 - 1;
    this.pointer.y = -((e.clientY - rect.top) / rect.height) * 2 + 1;

    this.raycaster.setFromCamera(this.pointer, this.camera);
    const wallMeshes = this.getWallMeshes();
    const hits = this.raycaster.intersectObjects(wallMeshes, false);

    if (hits.length > 0) {
      const hit = hits[0].object as THREE.Mesh;
      if (this.placementMode !== 'none') {
        this.startCutoutPlacement(hit);
      } else {
        this.selectWall(hit);
      }
      // Prevent orbit controls consuming this event
      e.stopPropagation();
    } else {
      // Clicked empty space — deselect
      if (this.placementMode !== 'none') {
        this.cancelCutoutPlacement();
      } else {
        this.deselectWall();
      }
    }
  }

  private onCanvasHover(e: PointerEvent) {
    const canvas = this.renderer.domElement;
    const rect = canvas.getBoundingClientRect();
    this.pointer.x = ((e.clientX - rect.left) / rect.width) * 2 - 1;
    this.pointer.y = -((e.clientY - rect.top) / rect.height) * 2 + 1;

    this.raycaster.setFromCamera(this.pointer, this.camera);
    const hits = this.raycaster.intersectObjects(this.getWallMeshes(), false);
    document.body.classList.toggle('wall-hover', hits.length > 0);
  }

  private selectWall(mesh: THREE.Mesh) {
    // Deselect previous
    if (this.selectedWall && this.selectedWall !== mesh) {
      this.dehighlightWall(this.selectedWall);
    }

    this.selectedWall = mesh;
    this.highlightWall(mesh);

    const entry = this.wallRegistry.get(mesh)!;
    this.openModal(entry, mesh);
  }

  private deselectWall() {
    if (this.selectedWall) {
      this.dehighlightWall(this.selectedWall);
      this.selectedWall = null;
    }
  }

  private highlightWall(mesh: THREE.Mesh) {
    const mat = mesh.material as THREE.MeshStandardMaterial;
    mat.emissive.set(0xf59e0b);
    mat.emissiveIntensity = 0.45;
  }

  private dehighlightWall(mesh: THREE.Mesh) {
    const mat = mesh.material as THREE.MeshStandardMaterial;
    mat.emissive.set(0x000000);
    mat.emissiveIntensity = 0;
  }

  // ─── Modal ────────────────────────────────────────────────────────────────
  private openModal(entry: WallEntry, mesh: THREE.Mesh) {
    const modal = document.getElementById('wall-modal')!;
    const labelEl = document.getElementById('wall-modal-label')!;
    const typeEl = document.getElementById('wall-modal-type')!;
    const heightIn = document.getElementById('wall-height-input') as HTMLInputElement;
    const lengthIn = document.getElementById('wall-length-input') as HTMLInputElement;
    const widthIn = document.getElementById('wall-width-input') as HTMLInputElement;
    const infoEl = document.getElementById('wall-modal-info')!;
    const mat = mesh.material as THREE.MeshStandardMaterial;

    labelEl.textContent = entry.label.replace(/(^\w+\s+){2}/, '') || entry.label;
    typeEl.textContent = entry.wallType;
    heightIn.value = entry.height.toFixed(2);
    lengthIn.value = entry.length.toFixed(2);
    widthIn.value = entry.thickness.toFixed(2);
    this.setModalColorControls(`#${mat.color.getHexString()}`);
    infoEl.textContent = `Floor: ${entry.record.settings.floor_level || '—'}`;
    infoEl.style.color = '';

    modal.style.display = 'flex';
    this.controls.enabled = false;
  }

  private closeModal() {
    const modal = document.getElementById('wall-modal')!;
    modal.style.display = 'none';
    this.deselectWall();
    this.controls.enabled = true;
  }

  private applyWallEdit() {
    if (!this.selectedWall) return;

    const mesh = this.selectedWall;
    const entry = this.wallRegistry.get(mesh)!;
    const heightIn = document.getElementById('wall-height-input') as HTMLInputElement;
    const lengthIn = document.getElementById('wall-length-input') as HTMLInputElement;
    const widthIn = document.getElementById('wall-width-input') as HTMLInputElement;
    const colorIn = document.getElementById('wall-color-input') as HTMLInputElement;

    const newHeight = parseFloat(heightIn.value);
    const newLength = parseFloat(lengthIn.value);
    const newThickness = parseFloat(widthIn.value);
    const newColorHex = this.normalizeHexColor(colorIn.value);

    if (
      isNaN(newHeight) || newHeight <= 0 ||
      isNaN(newLength) || newLength <= 0 ||
      isNaN(newThickness) || newThickness <= 0 ||
      !newColorHex
    ) return;

    // Rebuild geometry with new dimensions
    mesh.geometry.dispose();
    mesh.geometry = new THREE.BoxGeometry(newLength, newHeight, newThickness);

    // Keep the wall centred at its original world-space midpoint,
    // just update the Y so it sits correctly for the new height
    mesh.position.copy(entry.worldPos);
    mesh.position.y = entry.baseElev + newHeight / 2;
    mesh.rotation.y = entry.worldRotY;

    const mat = mesh.material as THREE.MeshStandardMaterial;
    mat.color.set(newColorHex);

    // Persist updated values back into the entry
    entry.height = newHeight;
    entry.length = newLength;
    entry.thickness = newThickness;
    entry.originalColor = mat.color.getHex();
    (entry.record.settings as any).color = newColorHex;
    // Update worldPos Y to match the new height centre
    entry.worldPos.y = mesh.position.y;
    this.addAutoFloorFromWalls();
    this.refreshEstimateIfOpen();

    // Close modal and deselect
    this.closeModal();
  }

  // ─── Scene helpers ────────────────────────────────────────────────────────
  private clearScene() {
    this.wallRegistry.clear();
    this.selectedWall = null;
    this.closeEstimateModal();
    this.updateEstimateButtonState();
    document.body.classList.remove('wall-hover');

    while (this.buildingGroup.children.length > 0) {
      const child = this.buildingGroup.children[0] as THREE.Mesh;
      if (child.geometry) child.geometry.dispose();
      if ((child as THREE.Mesh).material) {
        const m = (child as THREE.Mesh).material;
        if (Array.isArray(m)) m.forEach(x => x.dispose());
        else (m as THREE.Material).dispose();
      }
      this.buildingGroup.remove(child);
    }
  }

  private addAutoFloorFromWalls() {
    const existingAutoFloor = this.buildingGroup.getObjectByName('auto-generated-floor');
    if (existingAutoFloor) {
      const mesh = existingAutoFloor as THREE.Mesh;
      if (mesh.geometry) mesh.geometry.dispose();
      const mat = mesh.material;
      if (Array.isArray(mat)) mat.forEach((m) => m.dispose());
      else if (mat) (mat as THREE.Material).dispose();
      this.buildingGroup.remove(mesh);
    }

    const wallMeshes = this.getWallMeshes();
    if (!wallMeshes.length) return;

    const globalBox = new THREE.Box3();
    let hasBounds = false;
    wallMeshes.forEach((wallMesh) => {
      const b = new THREE.Box3().setFromObject(wallMesh);
      if (!Number.isFinite(b.min.x) || !Number.isFinite(b.max.x)) return;
      if (!hasBounds) {
        globalBox.copy(b);
        hasBounds = true;
      } else {
        globalBox.union(b);
      }
    });
    if (!hasBounds) return;

    const width = Math.max(0.5, globalBox.max.x - globalBox.min.x + 0.6);
    const depth = Math.max(0.5, globalBox.max.z - globalBox.min.z + 0.6);
    const floorThickness = 0.25;
    const topY = globalBox.min.y + 0.01;

    const floorGeo = new THREE.BoxGeometry(width, floorThickness, depth);
    const floorMat = new THREE.MeshStandardMaterial({
      color: 0x0f172a,
      roughness: 0.88,
      metalness: 0.02,
      transparent: true,
      opacity: 0.90,
    });
    const floorMesh = new THREE.Mesh(floorGeo, floorMat);
    floorMesh.name = 'auto-generated-floor';
    floorMesh.position.set(
      (globalBox.min.x + globalBox.max.x) / 2,
      topY - floorThickness / 2,
      (globalBox.min.z + globalBox.max.z) / 2
    );
    floorMesh.receiveShadow = true;
    floorMesh.castShadow = false;
    this.buildingGroup.add(floorMesh);
  }

  private frameCamera() {
    const box = new THREE.Box3().setFromObject(this.buildingGroup);
    const size = box.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z, 10);
    const dist = maxDim * 1.8;
    // Position camera relative to the object's center so framing works for tall buildings
    const center = box.getCenter(new THREE.Vector3());
    this.camera.position.set(center.x + dist, center.y + dist * 0.7, center.z + dist);
    this.controls.target.copy(center);
    this.controls.update();
  }

  private updateLegendForNewFormat() {
    const legend = document.querySelector('.legend') as HTMLElement;
    legend.innerHTML = Object.entries(LABEL_MAP).map(([type, label]) => {
      const hex = (MATERIAL_COLORS[type] ?? MATERIAL_COLORS.default).toString(16).padStart(6, '0');
      return `<div class="legend-item"><span class="color" style="background:#${hex}"></span>${label}</div>`;
    }).join('');
  }

  private updateLegendForOldFormat() {
    const legend = document.querySelector('.legend') as HTMLElement;
    // Remove hard-coded wall-type legend for old-format input so the UI
    // doesn't always show supported wall types. Leave legend empty by default.
    legend.innerHTML = '';
  }

  private animate() {
    requestAnimationFrame(() => this.animate());
    this.controls.update();
    this.renderer.render(this.scene, this.camera);
  }
}

new HouseViewer();
