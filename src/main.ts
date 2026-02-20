import './style.css';
import * as THREE from 'three';
import { OrbitControls } from 'three/examples/jsm/controls/OrbitControls.js';
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
  };
  coordinates_real_world: [number, number][];
  scale_factor_float: number;
}
interface NewJsonData {
  project_id: string;
  records: NewRecord[];
}

// ─── SCALE helpers ───────────────────────────────────────────────────────────
// real_world coords are in inches at 1/4" plan scale → 1 three.js unit = 1 foot
const SCALE = 1 / 12; // inches → feet

const MATERIAL_COLORS: Record<string, number> = {
  roof_system: 0xf97316, // orange  – roof panel / rafter area
  eave_length: 0x22d3ee, // teal    – eave edges
  ridge_length: 0xa855f7, // purple  – ridge
  hip_length: 0xfbbf24, // amber   – hip
  valley_length: 0x60a5fa, // blue    – valley
  gable_length: 0x4ade80, // green   – gable
  default: 0xffffff,
};

const LABEL_MAP: Record<string, string> = {
  roof_system: 'Roof Panel / Rafter',
  eave_length: 'Eave',
  ridge_length: 'Ridge',
  hip_length: 'Hip',
  valley_length: 'Valley',
  gable_length: 'Gable',
};

// ─── ELEVATION helpers (so lines don't overlap) ───────────────────────────────
const ELEVATIONS: Record<string, number> = {
  roof_system: 0,
  eave_length: 0.05,
  valley_length: -0.05,
  hip_length: 0.15,
  ridge_length: 0.3,
  gable_length: 0.1,
  default: 0,
};

class HouseViewer {
  private scene: THREE.Scene;
  private camera: THREE.PerspectiveCamera;
  private renderer: THREE.WebGLRenderer;
  private controls: OrbitControls;
  private buildingGroup: THREE.Group;

  constructor() {
    this.scene = new THREE.Scene();
    this.scene.background = new THREE.Color(0x0f172a);

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
    this.animate();
  }

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

  private initEventListeners() {
    window.addEventListener('resize', () => {
      this.camera.aspect = window.innerWidth / window.innerHeight;
      this.camera.updateProjectionMatrix();
      this.renderer.setSize(window.innerWidth, window.innerHeight);
    });

    const fileInput = document.getElementById('json-upload') as HTMLInputElement;
    fileInput.addEventListener('change', (e) => this.handleFileUpload(e));
  }

  private async handleFileUpload(event: Event) {
    const target = event.target as HTMLInputElement;
    const file = target.files?.[0];
    if (!file) return;

    const fileInfo = document.querySelector('.file-info') as HTMLElement;
    fileInfo.textContent = `Loading: ${file.name}…`;

    try {
      const text = await file.text();
      const data = JSON.parse(text);

      // ─── format detection ───────────────────────────────────────────────
      if (data.records && Array.isArray(data.records)) {
        this.renderNewFormat(data as NewJsonData, file.name);
      } else if (data.walls && Array.isArray(data.walls)) {
        this.renderOldFormat(data as OldJsonData, file.name);
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

    // ── 1. find global bounds for centering ──────────────────────────────
    let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
    records.forEach(rec => {
      rec.coordinates_real_world.forEach(([x, y]) => {
        minX = Math.min(minX, x); maxX = Math.max(maxX, x);
        minY = Math.min(minY, y); maxY = Math.max(maxY, y);
      });
    });
    const cx = (minX + maxX) / 2;
    const cy = (minY + maxY) / 2;

    // ── 2. grouped stats ─────────────────────────────────────────────────
    const typeCounts: Record<string, number> = {};

    // ── 3. render each record ────────────────────────────────────────────
    records.forEach(rec => {
      const pts = rec.coordinates_real_world;
      if (!pts || pts.length === 0) return;

      const mt = rec.materialType || 'default';
      typeCounts[mt] = (typeCounts[mt] || 0) + 1;

      const color = MATERIAL_COLORS[mt] ?? MATERIAL_COLORS.default;
      const elev = ELEVATIONS[mt] ?? ELEVATIONS.default;
      const pitch = parseFloat(rec.settings.pitch || '4');

      // convert to three.js coords (inches → feet, flip Y for plan orientation)
      const toV3 = ([x, y]: [number, number], z = 0): THREE.Vector3 =>
        new THREE.Vector3((x - cx) * SCALE, z, (y - cy) * SCALE);

      if (pts.length >= 3) {
        // ─ POLYGON: extruded roof panel ──────────────────────────────
        this.renderPolygon(pts, cx, cy, pitch, color, elev);
      } else if (pts.length === 2) {
        // ─ LINE: eave / ridge / hip / valley ────────────────────────
        this.renderLine(toV3(pts[0], elev), toV3(pts[1], elev), color, mt);
      }
    });

    // ── 4. update stats panel ─────────────────────────────────────────────
    const statsPanel = document.getElementById('stats-panel') as HTMLElement;
    statsPanel.style.display = 'block';

    const statsList = document.getElementById('stats-list') as HTMLElement;
    statsList.innerHTML = '';
    Object.entries(typeCounts).forEach(([type, count]) => {
      const div = document.createElement('div');
      div.className = 'stat-row';
      const dot = document.createElement('span');
      dot.className = 'stat-dot';
      dot.style.background = '#' + (MATERIAL_COLORS[type] ?? MATERIAL_COLORS.default).toString(16).padStart(6, '0');
      div.appendChild(dot);
      div.appendChild(document.createTextNode(`${LABEL_MAP[type] ?? type}: ${count}`));
      statsList.appendChild(div);
    });

    const fileInfo = document.querySelector('.file-info') as HTMLElement;
    fileInfo.textContent = `Project ${data.project_id} — ${records.length} records loaded`;

    this.frameCamera();
    this.updateLegendForNewFormat();
  }

  private renderPolygon(
    pts: [number, number][],
    cx: number,
    cy: number,
    _pitch: number,
    color: number,
    baseElev: number
  ) {
    // Use a ShapeGeometry for the footprint, then apply pitch as a slight X-rotation
    const shape = new THREE.Shape();
    const first = pts[0];
    shape.moveTo((first[0] - cx) * SCALE, (first[1] - cy) * SCALE);
    for (let i = 1; i < pts.length; i++) {
      shape.lineTo((pts[i][0] - cx) * SCALE, (pts[i][1] - cy) * SCALE);
    }
    shape.closePath();

    const geo = new THREE.ShapeGeometry(shape);
    const mat = new THREE.MeshStandardMaterial({
      color,
      metalness: 0.05,
      roughness: 0.7,
      side: THREE.DoubleSide,
      transparent: true,
      opacity: 0.82,
    });

    const mesh = new THREE.Mesh(geo, mat);
    mesh.rotation.x = -Math.PI / 2; // lay flat in XZ plane
    mesh.position.y = baseElev;
    this.buildingGroup.add(mesh);

    // Wireframe border
    const edgesGeo = new THREE.EdgesGeometry(geo);
    const edgesMat = new THREE.LineBasicMaterial({ color: 0xffffff, opacity: 0.3, transparent: true });
    const edges = new THREE.LineSegments(edgesGeo, edgesMat);
    edges.rotation.x = -Math.PI / 2;
    edges.position.y = baseElev + 0.01;
    this.buildingGroup.add(edges);
  }

  private renderLine(
    a: THREE.Vector3,
    b: THREE.Vector3,
    color: number,
    type: string
  ) {
    const dir = new THREE.Vector3().subVectors(b, a);
    const length = dir.length();
    if (length < 0.001) return;

    // draw as a thin flat box so it has visible thickness from above
    const thickness = type === 'ridge_length' ? 0.25 : 0.15;
    const height = type === 'ridge_length' ? 0.3 : 0.12;

    const geo = new THREE.BoxGeometry(length, height, thickness);
    const mat = new THREE.MeshStandardMaterial({ color, metalness: 0.2, roughness: 0.6 });
    const mesh = new THREE.Mesh(geo, mat);

    // position at midpoint between a and b
    mesh.position.copy(a).lerp(b, 0.5);
    mesh.position.y += height / 2;

    // rotate to align with direction
    const angle = Math.atan2(dir.z, dir.x);
    mesh.rotation.y = -angle;

    this.buildingGroup.add(mesh);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // OLD FORMAT renderer (kept for backwards compatibility)
  // ═══════════════════════════════════════════════════════════════════════════
  private wallMaterials: Record<string, THREE.Material> = {
    perimeter_wall: new THREE.MeshStandardMaterial({ color: 0xef4444, metalness: 0.1, roughness: 0.8 }),
    interior_wall: new THREE.MeshStandardMaterial({ color: 0x3b82f6, metalness: 0.1, roughness: 0.8 }),
    foundation_wall: new THREE.MeshStandardMaterial({ color: 0x64748b, metalness: 0.1, roughness: 0.8 }),
    knee_wall: new THREE.MeshStandardMaterial({ color: 0xf59e0b, metalness: 0.1, roughness: 0.8 }),
    default: new THREE.MeshStandardMaterial({ color: 0xffffff, metalness: 0.1, roughness: 0.8 }),
  };

  private renderOldFormat(data: OldJsonData, _fileName: string) {
    this.clearScene();

    const walls = data.walls;
    if (!walls || walls.length === 0) return;

    let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
    walls.forEach(w => {
      minX = Math.min(minX, w.bbox.x1, w.bbox.x2);
      maxX = Math.max(maxX, w.bbox.x1, w.bbox.x2);
      minY = Math.min(minY, w.bbox.y1, w.bbox.y2);
      maxY = Math.max(maxY, w.bbox.y1, w.bbox.y2);
    });

    const cx = (minX + maxX) / 2;
    const cy = (minY + maxY) / 2;
    const size = Math.max(maxX - minX, maxY - minY);
    const sc = 50 / size;
    const wallH = 2.5;

    walls.forEach(w => {
      const width = Math.abs(w.bbox.x2 - w.bbox.x1) * sc || 0.1;
      const depth = Math.abs(w.bbox.y2 - w.bbox.y1) * sc || 0.1;
      const geo = new THREE.BoxGeometry(width, wallH, depth);
      const mat = this.wallMaterials[w.class] || this.wallMaterials.default;
      const mesh = new THREE.Mesh(geo, mat as THREE.Material);
      mesh.position.set((w.center.x - cx) * sc, wallH / 2, (w.center.y - cy) * sc);
      this.buildingGroup.add(mesh);
    });

    const statsPanel = document.getElementById('stats-panel') as HTMLElement;
    statsPanel.style.display = 'block';
    const statsList = document.getElementById('stats-list') as HTMLElement;
    statsList.innerHTML = `<div class="stat-row">Walls: ${walls.length}</div>`;

    const fileInfo = document.querySelector('.file-info') as HTMLElement;
    fileInfo.textContent = `Applied: ${_fileName}`;

    this.frameCamera();
    this.updateLegendForOldFormat();
  }

  // ─── helpers ─────────────────────────────────────────────────────────────
  private clearScene() {
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

  private frameCamera() {
    const box = new THREE.Box3().setFromObject(this.buildingGroup);
    const size = box.getSize(new THREE.Vector3());
    const maxDim = Math.max(size.x, size.y, size.z, 10);
    const dist = maxDim * 1.8;
    this.camera.position.set(dist, dist * 0.7, dist);
    this.controls.target.set(0, 0, 0);
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
    legend.innerHTML = `
      <div class="legend-item"><span class="color perimeter"></span>Perimeter Wall</div>
      <div class="legend-item"><span class="color interior"></span>Interior Wall</div>
      <div class="legend-item"><span class="color foundation"></span>Foundation</div>
      <div class="legend-item"><span class="color knee"></span>Knee Wall</div>
    `;
  }

  private animate() {
    requestAnimationFrame(() => this.animate());
    this.controls.update();
    this.renderer.render(this.scene, this.camera);
  }
}

new HouseViewer();
