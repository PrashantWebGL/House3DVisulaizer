import './style.css';
import * as THREE from 'three';
import { OrbitControls } from 'three/examples/jsm/controls/OrbitControls.js';
import { PointerLockControls } from 'three/examples/jsm/controls/PointerLockControls.js';
import { CSG } from 'three-csg-ts';
import { GLTFExporter } from 'three/examples/jsm/exporters/GLTFExporter.js';
import { inject } from '@vercel/analytics';
import defaultHouseJson from '../assets/Small_houseClean.json';

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

// ─── Page-format (Plan) JSON ───────────────────────────────────────────────
interface PlanWallCoordinates {
  x1: number; y1: number; x2: number; y2: number;
}

interface PlanWall {
  geometry?: { coordinates?: PlanWallCoordinates };
  properties?: { thickness_inches?: number; wall_height?: string | number; category?: string; floor_label?: string };
  category?: string;
}

interface PlanEntities {
  walls?: PlanWall[];
  windows_and_doors_floor_plans?: any;
  roofing?: any;
}

interface PlanPage {
  page_index?: number;
  source_file_id?: string;
  page_number_in_source_file?: number;
  entities?: PlanEntities;
}

interface PlanJson {
  pages?: PlanPage[];
  project_metadata?: Record<string, unknown>;
  project_id?: string | number;
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

interface WallEditSnapshot {
  height: number;
  length: number;
  thickness: number;
  colorHex: string;
  textureKey: string;
  baseElev: number;
  worldPos: THREE.Vector3;
  worldRotY: number;
}

interface WallEditHistoryEntry {
  mesh: THREE.Mesh;
  before: WallEditSnapshot;
  after: WallEditSnapshot;
}

class WaypointManager {
  private readonly eyeHeight: number;
  private autoPoints: THREE.Vector3[] = [];
  private customPoints: THREE.Vector3[] = [];

  constructor(eyeHeight = 5.5) {
    this.eyeHeight = eyeHeight;
  }

  public rebuildFromWalls(wallMeshes: THREE.Mesh[]) {
    const box = new THREE.Box3();
    let hasBounds = false;
    wallMeshes.forEach((mesh) => {
      const b = new THREE.Box3().setFromObject(mesh);
      if (!Number.isFinite(b.min.x) || !Number.isFinite(b.max.x)) return;
      if (!hasBounds) {
        box.copy(b);
        hasBounds = true;
      } else {
        box.union(b);
      }
    });

    if (!hasBounds) {
      this.autoPoints = [];
      return;
    }

    const minX = box.min.x;
    const maxX = box.max.x;
    const minZ = box.min.z;
    const maxZ = box.max.z;
    const marginX = Math.max(1.2, (maxX - minX) * 0.12);
    const marginZ = Math.max(1.2, (maxZ - minZ) * 0.12);
    const centerX = (minX + maxX) / 2;
    const centerZ = (minZ + maxZ) / 2;

    this.autoPoints = [
      new THREE.Vector3(minX + marginX, this.eyeHeight, maxZ - marginZ), // entrance
      new THREE.Vector3(centerX, this.eyeHeight, centerZ),                 // living
      new THREE.Vector3(maxX - marginX, this.eyeHeight, centerZ),          // kitchen
      new THREE.Vector3(maxX - marginX, this.eyeHeight, minZ + marginZ),   // bedroom 1
      new THREE.Vector3(minX + marginX, this.eyeHeight, minZ + marginZ),   // bedroom 2
    ];
  }

  public addCustomStop(point: THREE.Vector3) {
    const p = point.clone();
    p.y = this.eyeHeight;
    this.customPoints.push(p);
  }

  public clearCustomStops() {
    this.customPoints = [];
  }

  public getCustomStops(): THREE.Vector3[] {
    return this.customPoints.map((p) => p.clone());
  }

  public getWaypoints(): THREE.Vector3[] {
    if (this.customPoints.length >= 2) return this.customPoints.map((p) => p.clone());
    return this.autoPoints.map((p) => p.clone());
  }
}

interface GuidedState {
  visible: boolean;
  current: number;
  total: number;
  canPrev: boolean;
  canNext: boolean;
}

class CameraMovement {
  private readonly camera: THREE.PerspectiveCamera;
  private readonly eyeHeight: number;
  private readonly velocity = new THREE.Vector3();
  private readonly input = { forward: false, backward: false, left: false, right: false };
  private active = false;

  private readonly maxSpeed = 12; // feet/sec
  private readonly accel = 12;
  private readonly damping = 10;

  constructor(camera: THREE.PerspectiveCamera, _controls: PointerLockControls, eyeHeight = 5.5) {
    this.camera = camera;
    this.eyeHeight = eyeHeight;
    document.addEventListener('keydown', this.onKeyDown);
    document.addEventListener('keyup', this.onKeyUp);
  }

  public setActive(isActive: boolean) {
    this.active = isActive;
    if (!isActive) this.velocity.set(0, 0, 0);
  }

  public update(delta: number, isBlocked: (candidate: THREE.Vector3) => boolean) {
    if (!this.active) return;

    const forward = new THREE.Vector3();
    this.camera.getWorldDirection(forward);
    forward.y = 0;
    if (forward.lengthSq() > 0) forward.normalize();
    const right = new THREE.Vector3(-forward.z, 0, forward.x);

    const desiredDir = new THREE.Vector3();
    if (this.input.forward) desiredDir.add(forward);
    if (this.input.backward) desiredDir.sub(forward);
    if (this.input.left) desiredDir.sub(right);
    if (this.input.right) desiredDir.add(right);
    if (desiredDir.lengthSq() > 0) desiredDir.normalize();

    const targetVel = desiredDir.multiplyScalar(this.maxSpeed);
    const blend = 1 - Math.exp(-this.accel * delta);
    this.velocity.lerp(targetVel, blend);
    if (desiredDir.lengthSq() === 0) {
      const decay = Math.exp(-this.damping * delta);
      this.velocity.multiplyScalar(decay);
    }

    const current = this.camera.position.clone();
    current.y = this.eyeHeight;

    const candidate = current.clone().addScaledVector(this.velocity, delta);
    candidate.y = this.eyeHeight;

    if (!isBlocked(candidate)) {
      this.camera.position.copy(candidate);
      return;
    }

    const tryX = current.clone();
    tryX.x = candidate.x;
    tryX.y = this.eyeHeight;

    const tryZ = current.clone();
    tryZ.z = candidate.z;
    tryZ.y = this.eyeHeight;

    if (!isBlocked(tryX)) {
      this.camera.position.copy(tryX);
      this.velocity.z = 0;
    } else if (!isBlocked(tryZ)) {
      this.camera.position.copy(tryZ);
      this.velocity.x = 0;
    } else {
      this.velocity.set(0, 0, 0);
    }
  }

  private onKeyDown = (e: KeyboardEvent) => {
    if (!this.active) return;
    if (e.code === 'KeyW') this.input.forward = true;
    if (e.code === 'KeyS') this.input.backward = true;
    if (e.code === 'KeyA') this.input.left = true;
    if (e.code === 'KeyD') this.input.right = true;
  };

  private onKeyUp = (e: KeyboardEvent) => {
    if (e.code === 'KeyW') this.input.forward = false;
    if (e.code === 'KeyS') this.input.backward = false;
    if (e.code === 'KeyA') this.input.left = false;
    if (e.code === 'KeyD') this.input.right = false;
  };
}

class WalkthroughController {
  private readonly camera: THREE.PerspectiveCamera;
  private readonly orbitControls: OrbitControls;
  private readonly pointerControls: PointerLockControls;
  private readonly movement: CameraMovement;
  private readonly waypointManager: WaypointManager;
  private readonly getWallMeshes: () => THREE.Mesh[];
  private readonly getDoorMeshes: () => THREE.Mesh[];
  private readonly getFloorMeshes: () => THREE.Mesh[];
  private readonly onGuidedStateChange: (state: GuidedState) => void;
  private readonly eyeHeight = 5.5;
  private collisionBoxes: THREE.Box3[] = [];
  private doorBoxes: THREE.Box3[] = [];
  private readonly markerGroup = new THREE.Group();

  private mode: 'none' | 'guided' | 'free' = 'none';
  private pathEditEnabled = false;
  private curve: THREE.CatmullRomCurve3 | null = null;
  private guidedPoints: THREE.Vector3[] = [];
  private guidedStopTs: number[] = [];
  private currentStopIndex = 0;
  private targetStopIndex = 0;
  private segmentMoving = false;
  private segmentProgress = 0;
  private segmentStartT = 0;
  private segmentEndT = 0;
  private segmentDuration = 0.1;
  private readonly guidedSpeed = 8; // feet/sec
  private readonly orbitDefaults: { rotate: boolean; zoom: boolean; pan: boolean };
  private guidedYaw = 0;
  private guidedPitch = 0;
  private guidedMouseDown = false;
  private lastMouseX = 0;
  private lastMouseY = 0;

  constructor(
    scene: THREE.Scene,
    camera: THREE.PerspectiveCamera,
    orbitControls: OrbitControls,
    domElement: HTMLElement,
    getWallMeshes: () => THREE.Mesh[],
    getDoorMeshes: () => THREE.Mesh[],
    getFloorMeshes: () => THREE.Mesh[],
    onGuidedStateChange: (state: GuidedState) => void
  ) {
    this.camera = camera;
    this.orbitControls = orbitControls;
    this.pointerControls = new PointerLockControls(camera, domElement);
    this.movement = new CameraMovement(camera, this.pointerControls, this.eyeHeight);
    this.waypointManager = new WaypointManager(this.eyeHeight);
    this.getWallMeshes = getWallMeshes;
    this.getDoorMeshes = getDoorMeshes;
    this.getFloorMeshes = getFloorMeshes;
    this.onGuidedStateChange = onGuidedStateChange;
    this.orbitDefaults = {
      rotate: orbitControls.enableRotate,
      zoom: orbitControls.enableZoom,
      pan: orbitControls.enablePan,
    };
    scene.add(this.pointerControls.getObject());
    scene.add(this.markerGroup);
  }

  public syncEnvironment() {
    const walls = this.getWallMeshes();
    this.waypointManager.rebuildFromWalls(walls);
    const radius = 0.38;
    this.collisionBoxes = walls.map((mesh) => {
      const box = new THREE.Box3().setFromObject(mesh);
      box.min.x -= radius;
      box.max.x += radius;
      box.min.z -= radius;
      box.max.z += radius;
      return box;
    });

    this.doorBoxes = this.getDoorMeshes().map((mesh) => {
      const b = new THREE.Box3().setFromObject(mesh);
      b.min.x -= 0.85;
      b.max.x += 0.85;
      b.min.z -= 0.85;
      b.max.z += 0.85;
      b.min.y = -Infinity;
      b.max.y = Infinity;
      return b;
    });
    this.updateStopMarkers();
  }

  public start(mode: 'guided' | 'free'): boolean {
    this.syncEnvironment();
    if (!this.getWallMeshes().length) return false;
    this.stop();
    this.mode = mode;
    this.orbitControls.enabled = false;
    this.orbitControls.enableRotate = false;
    this.orbitControls.enableZoom = false;
    this.orbitControls.enablePan = false;

    if (mode === 'guided') return this.startGuided();
    return this.startFree();
  }

  public stop() {
    if (this.mode === 'free' && this.pointerControls.isLocked) {
      this.pointerControls.unlock();
    }
    this.movement.setActive(false);
    this.mode = 'none';
    this.curve = null;
    this.guidedPoints = [];
    this.guidedStopTs = [];
    this.currentStopIndex = 0;
    this.targetStopIndex = 0;
    this.segmentMoving = false;
    this.segmentProgress = 0;
    this.orbitControls.enabled = true;
    this.orbitControls.enableRotate = this.orbitDefaults.rotate;
    this.orbitControls.enableZoom = this.orbitDefaults.zoom;
    this.orbitControls.enablePan = this.orbitDefaults.pan;
    this.emitGuidedState();
  }

  public handleGuidedPointerDown(e: PointerEvent): boolean {
    if (this.mode !== 'guided') return false;
    this.guidedMouseDown = true;
    this.lastMouseX = e.clientX;
    this.lastMouseY = e.clientY;
    return true;
  }

  public handleGuidedPointerMove(e: PointerEvent): boolean {
    if (this.mode !== 'guided' || !this.guidedMouseDown) return false;
    const dx = e.clientX - this.lastMouseX;
    const dy = e.clientY - this.lastMouseY;
    this.lastMouseX = e.clientX;
    this.lastMouseY = e.clientY;

    this.guidedYaw -= dx * 0.003;
    this.guidedPitch -= dy * 0.003;
    this.guidedPitch = Math.max(-1.2, Math.min(1.2, this.guidedPitch));
    return true;
  }

  public handleGuidedPointerUp(): boolean {
    if (this.mode !== 'guided') return false;
    this.guidedMouseDown = false;
    return true;
  }

  public handleGuidedWheel(e: WheelEvent): boolean {
    if (this.mode !== 'guided') return false;
    const nextFov = this.camera.fov + (e.deltaY * 0.02);
    this.camera.fov = Math.max(35, Math.min(90, nextFov));
    this.camera.updateProjectionMatrix();
    return true;
  }

  public update(delta: number) {
    if (this.mode === 'guided' && this.curve) {
      if (this.segmentMoving) {
        this.segmentProgress = Math.min(1, this.segmentProgress + (delta / this.segmentDuration));
        const t = this.segmentStartT + ((this.segmentEndT - this.segmentStartT) * this.segmentProgress);
        this.positionCameraOnCurve(t, this.segmentEndT >= this.segmentStartT ? 1 : -1);

        if (this.segmentProgress >= 1) {
          this.segmentMoving = false;
          this.currentStopIndex = this.targetStopIndex;
          this.emitGuidedState();
        }
      } else {
        this.positionCameraOnCurve(this.guidedStopTs[this.currentStopIndex], 1);
      }
      return;
    }

    if (this.mode === 'free') {
      this.camera.position.y = this.eyeHeight;
      this.movement.update(delta, (candidate) => this.isBlocked(candidate));
    }
  }

  public isActive(): boolean {
    return this.mode !== 'none';
  }

  public setPathEditEnabled(enabled: boolean) {
    this.pathEditEnabled = enabled;
  }

  public clearCustomStops() {
    this.waypointManager.clearCustomStops();
    this.updateStopMarkers();
    this.emitGuidedState();
  }

  public tryAddStopFromRay(raycaster: THREE.Raycaster): boolean {
    if (!this.pathEditEnabled) return false;
    const floorMeshes = this.getFloorMeshes();
    if (!floorMeshes.length) return false;

    const hits = raycaster.intersectObjects(floorMeshes, false);
    if (!hits.length) return false;

    const p = hits[0].point.clone();
    p.y = this.eyeHeight;
    this.waypointManager.addCustomStop(p);
    this.updateStopMarkers();
    this.emitGuidedState();
    return true;
  }

  public goNextStop() {
    if (this.mode !== 'guided' || this.segmentMoving) return;
    if (this.currentStopIndex >= this.guidedPoints.length - 1) return;
    this.startSegment(this.currentStopIndex + 1);
  }

  public goPrevStop() {
    if (this.mode !== 'guided' || this.segmentMoving) return;
    if (this.currentStopIndex <= 0) return;
    this.startSegment(this.currentStopIndex - 1);
  }

  private startGuided(): boolean {
    const points = this.waypointManager.getWaypoints();
    if (points.length < 2) {
      this.stop();
      return false;
    }
    this.guidedPoints = points.map((p) => p.clone().setY(this.eyeHeight));
    this.curve = new THREE.CatmullRomCurve3(this.guidedPoints, false, 'centripetal', 0.35);
    this.curve.arcLengthDivisions = 200;
    this.guidedStopTs = this.computeStopTs(this.guidedPoints, this.curve);
    this.currentStopIndex = 0;
    this.targetStopIndex = 0;
    this.segmentMoving = false;
    this.segmentProgress = 0;
    this.guidedYaw = 0;
    this.guidedPitch = 0;
    this.guidedMouseDown = false;
    this.camera.position.copy(this.guidedPoints[0]);
    this.camera.lookAt(this.guidedPoints[1]);
    this.emitGuidedState();
    return true;
  }

  private startFree(): boolean {
    const points = this.waypointManager.getWaypoints();
    if (points.length > 0) this.camera.position.copy(points[0]);
    this.camera.position.y = this.eyeHeight;
    this.movement.setActive(true);
    this.pointerControls.lock();
    this.emitGuidedState();
    return true;
  }

  private startSegment(targetStopIndex: number) {
    if (!this.curve) return;
    this.targetStopIndex = targetStopIndex;
    this.segmentStartT = this.guidedStopTs[this.currentStopIndex];
    this.segmentEndT = this.guidedStopTs[targetStopIndex];
    const segLength = this.sampleCurveDistance(this.segmentStartT, this.segmentEndT);
    this.segmentDuration = Math.max(0.25, segLength / this.guidedSpeed);
    this.segmentProgress = 0;
    this.segmentMoving = true;
    this.emitGuidedState();
  }

  private sampleCurveDistance(startT: number, endT: number): number {
    if (!this.curve) return 0;
    const steps = 48;
    let distance = 0;
    let prev = this.curve.getPointAt(startT);
    for (let i = 1; i <= steps; i++) {
      const t = startT + ((endT - startT) * (i / steps));
      const p = this.curve.getPointAt(t);
      distance += p.distanceTo(prev);
      prev = p;
    }
    return distance;
  }

  private positionCameraOnCurve(t: number, direction: 1 | -1) {
    if (!this.curve) return;
    const clampedT = Math.max(0, Math.min(1, t));
    const point = this.curve.getPointAt(clampedT);
    const lookT = Math.max(0, Math.min(1, clampedT + (direction * 0.003)));
    const tangentPoint = this.curve.getPointAt(lookT);
    this.camera.position.set(point.x, this.eyeHeight, point.z);
    const baseForward = tangentPoint.clone().sub(this.camera.position).normalize();
    const lookOffset = new THREE.Vector3(0, 0, -1).applyEuler(new THREE.Euler(this.guidedPitch, this.guidedYaw, 0, 'YXZ'));
    const mixedForward = baseForward.clone().add(lookOffset.multiplyScalar(0.85)).normalize();
    const lookTarget = this.camera.position.clone().add(mixedForward);
    lookTarget.y = this.camera.position.y + mixedForward.y;
    this.camera.lookAt(lookTarget);
  }

  private computeStopTs(points: THREE.Vector3[], curve: THREE.CatmullRomCurve3): number[] {
    const ts: number[] = [];
    const samples = 600;
    const sampledPoints: THREE.Vector3[] = [];
    for (let i = 0; i <= samples; i++) {
      sampledPoints.push(curve.getPoint(i / samples));
    }

    points.forEach((stopPoint, idx) => {
      if (idx === 0) {
        ts.push(0);
        return;
      }
      if (idx === points.length - 1) {
        ts.push(1);
        return;
      }
      let bestI = 0;
      let bestD = Infinity;
      for (let i = 0; i <= samples; i++) {
        const d = sampledPoints[i].distanceTo(stopPoint);
        if (d < bestD) {
          bestD = d;
          bestI = i;
        }
      }
      ts.push(bestI / samples);
    });

    // Ensure increasing order for stable segment movement.
    for (let i = 1; i < ts.length; i++) {
      ts[i] = Math.max(ts[i], ts[i - 1]);
    }
    ts[ts.length - 1] = 1;
    return ts;
  }

  private updateStopMarkers() {
    while (this.markerGroup.children.length) {
      const c = this.markerGroup.children[0] as THREE.Mesh;
      if (c.geometry) c.geometry.dispose();
      const mat = c.material;
      if (Array.isArray(mat)) mat.forEach((m) => m.dispose());
      else if (mat) (mat as THREE.Material).dispose();
      this.markerGroup.remove(c);
    }

    const points = this.waypointManager.getCustomStops();
    points.forEach((point, idx) => {
      const marker = new THREE.Mesh(
        new THREE.SphereGeometry(0.23, 14, 12),
        new THREE.MeshStandardMaterial({
          color: idx === 0 ? 0x22c55e : 0x60a5fa,
          emissive: 0x0f172a,
          roughness: 0.25,
          metalness: 0.12,
        })
      );
      marker.position.set(point.x, 0.2, point.z);
      marker.userData.walkthroughMarker = true;
      this.markerGroup.add(marker);
    });
  }

  private emitGuidedState() {
    const total = this.guidedPoints.length;
    const visible = this.mode === 'guided' && total > 1;
    this.onGuidedStateChange({
      visible,
      current: visible ? this.currentStopIndex + 1 : 0,
      total: visible ? total : 0,
      canPrev: visible && !this.segmentMoving && this.currentStopIndex > 0,
      canNext: visible && !this.segmentMoving && this.currentStopIndex < total - 1,
    });
  }

  private isBlocked(candidate: THREE.Vector3): boolean {
    const nearDoorPortal = this.doorBoxes.some((doorBox) => doorBox.distanceToPoint(candidate) < 0.75);
    if (nearDoorPortal) return false;

    for (const box of this.collisionBoxes) {
      if (!box.containsPoint(candidate)) continue;
      const isInsideDoorPortal = this.doorBoxes.some((doorBox) => doorBox.containsPoint(candidate));
      if (isInsideDoorPortal) continue;
      return true;
    }
    return false;
  }
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

const WALL_TEXTURE_URLS: Record<string, string> = {
  animalgrey: new URL('../assets/WallTexture/animalgrey.jpg', import.meta.url).href,
  courtyard: new URL('../assets/WallTexture/courtyard.jpg', import.meta.url).href,
  geometric: new URL('../assets/WallTexture/geometric.jpg', import.meta.url).href,
  myrawhite: new URL('../assets/WallTexture/myrawhite.jpg', import.meta.url).href,
  plaingreen: new URL('../assets/WallTexture/plaingreen.jpg', import.meta.url).href,
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

// ─── Plan helpers ───────────────────────────────────────────────────────────
function parseFeetInches(raw: unknown, fallback = 9): number {
  if (typeof raw === 'number' && Number.isFinite(raw)) return raw;
  if (typeof raw !== 'string') return fallback;

  const text = raw.trim();
  if (!text) return fallback;

  // Extract feet and inches (with optional fractions)
  const feetMatch = text.match(/(-?\d+)\s*(?:'|ft)/i);
  const inchMatch = text.match(/(-?[\d\.]+(?:\s*\d+\/\d+)?)\s*(?:"|in)/i);

  let feet = feetMatch ? parseFloat(feetMatch[1]) : 0;

  let inches = 0;
  if (inchMatch) {
    const rawIn = inchMatch[1].replace(/\s+/g, '');
    if (rawIn.includes('/')) {
      const [whole, frac] = rawIn.split(/(?=\d+\/)/);
      const wholeIn = whole ? parseFloat(whole) : 0;
      const [num, den] = (frac || '0/1').split('/').map((v) => parseFloat(v));
      inches = wholeIn + (den ? num / den : 0);
    } else {
      inches = parseFloat(rawIn) || 0;
    }
  } else if (!feetMatch && /^-?\d+(\.\d+)?$/.test(text)) {
    // Plain number → feet
    feet = parseFloat(text);
  }

  return feet + inches / 12 || fallback;
}

function normalizePlanLabel(category?: string): string {
  if (!category) return 'Wall';
  const c = category.toLowerCase();
  if (c.includes('exterior')) return 'Exterior Wall';
  if (c.includes('interior')) return 'Interior Wall';
  if (c.includes('perimeter')) return 'Perimeter Wall';
  if (c.includes('foundation')) return 'Foundation Wall';
  if (c.includes('knee')) return 'Knee Wall';
  return 'Wall';
}

type OpeningBox = { xmin: number; xmax: number; ymin: number; ymax: number };

function toOpeningBox(box: any): OpeningBox | null {
  if (!box) return null;
  if (Array.isArray(box) && box.length === 4 && box.every((v) => typeof v === 'number')) {
    const [xmin, ymin, xmax, ymax] = box;
    if ([xmin, ymin, xmax, ymax].some((v) => !Number.isFinite(v))) return null;
    return { xmin, xmax, ymin, ymax };
  }

  if (Array.isArray(box) && box.length === 4 && Array.isArray(box[0])) {
    const xs: number[] = [];
    const ys: number[] = [];
    box.forEach((pt: any) => {
      if (Array.isArray(pt) && pt.length >= 2) {
        xs.push(Number(pt[0]));
        ys.push(Number(pt[1]));
      }
    });
    if (!xs.length || !ys.length) return null;
    if (!xs.every(Number.isFinite) || !ys.every(Number.isFinite)) return null;
    return { xmin: Math.min(...xs), xmax: Math.max(...xs), ymin: Math.min(...ys), ymax: Math.max(...ys) };
  }
  return null;
}

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
  private clock = new THREE.Clock();
  private walkthroughController: WalkthroughController;

  // ─── Wall selection ─────────────────────────────────────────────────────
  private raycaster = new THREE.Raycaster();
  private pointer = new THREE.Vector2();
  /** Maps every wall mesh → its data so we can edit / re-render it */
  private wallRegistry = new Map<THREE.Mesh, WallEntry>();
  private selectedWall: THREE.Mesh | null = null;
  private wallEditEnabled = false;
  private estimateRows: EstimateRow[] = [];
  private textureLoader = new THREE.TextureLoader();
  private wallTextureCache = new Map<string, THREE.Texture>();
  private activeModalTextureKey = 'none';
  private readonly maxHistorySize = 20;
  private undoStack: WallEditHistoryEntry[] = [];
  private redoStack: WallEditHistoryEntry[] = [];
  private sourceGroups = new Map<string, THREE.Mesh[]>();
  private sourceVisibility = new Map<string, boolean>();
  private pageGroups = new Map<string, { id: string; meshes: THREE.Mesh[]; roofMeshes: THREE.Mesh[]; sourceId: string; label: string }>();
  private pageVisibility = new Map<string, boolean>();
  private pageRoofVisibility = new Map<string, boolean>();
  private assemblyCollapsed = false;

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
    this.walkthroughController = new WalkthroughController(
      this.scene,
      this.camera,
      this.controls,
      this.renderer.domElement,
      () => this.getWallMeshes(),
      () => this.getDoorOpeningMeshes(),
      () => this.getFloorMeshes(),
      (state) => this.updateGuidedNavUI(state)
    );

    this.initLights();
    this.initGrid();
    this.initEventListeners();
    this.initModalListeners();
    this.initEstimateListeners();
    this.initAssemblyPanel();
    this.initExportListeners();
    this.loadDataFromJson(defaultHouseJson, 'Small_houseClean.json');
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
    canvas.addEventListener('pointermove', (e) => {
      if (this.walkthroughController.handleGuidedPointerMove(e)) {
        e.preventDefault();
      }
    });
    canvas.addEventListener('pointerup', (e) => {
      if (this.walkthroughController.handleGuidedPointerUp()) {
        e.preventDefault();
      }
    });
    canvas.addEventListener('wheel', (e) => {
      if (this.walkthroughController.handleGuidedWheel(e)) {
        e.preventDefault();
      }
    }, { passive: false });

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

    document.addEventListener('keydown', (e) => this.handleUndoRedoShortcuts(e));

    const startWalkBtn = document.getElementById('start-walkthrough-btn') as HTMLButtonElement;
    const stopWalkBtn = document.getElementById('stop-walkthrough-btn') as HTMLButtonElement;
    const modeSelect = document.getElementById('walkthrough-mode') as HTMLSelectElement;
    const guidedPathEditToggle = document.getElementById('guided-path-edit-toggle') as HTMLInputElement;
    const guidedClearStopsBtn = document.getElementById('guided-clear-stops-btn') as HTMLButtonElement;
    const guidedPrevBtn = document.getElementById('guided-prev-btn') as HTMLButtonElement;
    const guidedNextBtn = document.getElementById('guided-next-btn') as HTMLButtonElement;
    const wallEditToggle = document.getElementById('wall-edit-toggle') as HTMLInputElement;

    startWalkBtn.addEventListener('click', () => {
      const mode = (modeSelect.value === 'free' ? 'free' : 'guided') as 'guided' | 'free';
      this.walkthroughController.syncEnvironment();
      this.walkthroughController.start(mode);
    });

    stopWalkBtn.addEventListener('click', () => {
      this.walkthroughController.stop();
    });

    guidedPathEditToggle.checked = false;
    guidedPathEditToggle.addEventListener('change', () => {
      const enable = guidedPathEditToggle.checked && modeSelect.value === 'guided';
      this.walkthroughController.setPathEditEnabled(enable);
    });

    modeSelect.addEventListener('change', () => {
      const enable = guidedPathEditToggle.checked && modeSelect.value === 'guided';
      this.walkthroughController.setPathEditEnabled(enable);
    });

    guidedClearStopsBtn.addEventListener('click', () => {
      this.walkthroughController.clearCustomStops();
    });

    guidedPrevBtn.addEventListener('click', () => this.walkthroughController.goPrevStop());
    guidedNextBtn.addEventListener('click', () => this.walkthroughController.goNextStop());

    wallEditToggle.checked = false;
    this.wallEditEnabled = false;
    wallEditToggle.addEventListener('change', () => {
      this.wallEditEnabled = wallEditToggle.checked;
      if (!this.wallEditEnabled && this.selectedWall) this.closeModal();
    });
  }

  private initModalListeners() {
    const modal = document.getElementById('wall-modal')!;
    const card = document.getElementById('wall-modal-card')!;
    const btnClose = document.getElementById('wall-modal-close')!;
    const btnApply = document.getElementById('wall-modal-apply')!;
    const colorIn = document.getElementById('wall-color-input') as HTMLInputElement;
    const textureSwatches = document.querySelectorAll<HTMLButtonElement>('.texture-swatch');

    // Close on backdrop click (outside the card)
    modal.addEventListener('pointerdown', (e) => {
      if (e.target === modal) this.closeModal();
    });
    btnClose.addEventListener('click', () => this.closeModal());
    btnApply.addEventListener('click', () => this.applyWallEdit());

    // Prevent card clicks from closing modal (prevent event bubbling)
    card.addEventListener('pointerdown', (e) => e.stopPropagation());

    colorIn.addEventListener('input', () => this.setModalColorControls(colorIn.value));
    textureSwatches.forEach((btn) => {
      btn.addEventListener('click', () => {
        const key = btn.dataset.textureKey || 'none';
        this.setModalTextureSelection(key);
      });
    });
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

  private initExportListeners() {
    const btn = document.getElementById('export-glb-btn') as HTMLButtonElement | null;
    if (!btn) return;
    btn.addEventListener('click', () => this.exportGlb());
  }

  private renderAssemblyTree() {
    const body = document.getElementById('assembly-body') as HTMLElement | null;
    if (!body) return;

    if (this.sourceGroups.size === 0) {
      body.innerHTML = '<p class="assembly-empty">Load a plan-format JSON to see sources.</p>';
      return;
    }

    body.innerHTML = '';
    this.sourceGroups.forEach((meshes, sourceId) => {
      const item = document.createElement('div');
      item.className = 'assembly-item';

      const row = document.createElement('div');
      row.className = 'assembly-row';

      const toggleBtn = document.createElement('button');
      toggleBtn.type = 'button';
      toggleBtn.className = 'assembly-toggle';
      toggleBtn.dataset.sourceId = sourceId;
      toggleBtn.textContent = '▾';

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.dataset.sourceId = sourceId;
      checkbox.checked = meshes.every((m) => m.visible !== false);

      const label = document.createElement('span');
      label.className = 'assembly-label';
      label.textContent = sourceId;

      const count = document.createElement('span');
      count.className = 'assembly-count';
      count.textContent = `${meshes.length} wall${meshes.length === 1 ? '' : 's'}`;

      row.append(toggleBtn, checkbox, label, count);

      const children = document.createElement('div');
      children.className = 'assembly-children';
      children.dataset.children = sourceId;

      // Pages under this source
      const pages = Array.from(this.pageGroups.values()).filter((p) => p.sourceId === sourceId);
      pages.forEach((page) => {
        const pageRow = document.createElement('div');
        pageRow.className = 'assembly-row';
        pageRow.style.paddingLeft = '0.25rem';

        const pageToggle = document.createElement('button');
        pageToggle.type = 'button';
        pageToggle.className = 'assembly-page-toggle';
        pageToggle.dataset.pageId = page.id;
        pageToggle.textContent = '▸';

        const pageCheckbox = document.createElement('input');
        pageCheckbox.type = 'checkbox';
        pageCheckbox.dataset.pageId = page.id;
        pageCheckbox.dataset.sourceId = sourceId;
        const pageMeshes = page.meshes;
        const pageVis = this.pageVisibility.get(page.id);
        pageCheckbox.checked = pageVis !== false && pageMeshes.every((m) => m.visible !== false);
        const roofMeshes = page.roofMeshes || [];

        const pageLabel = document.createElement('span');
        pageLabel.className = 'assembly-label';
        pageLabel.textContent = page.label;

        const pageCount = document.createElement('span');
        pageCount.className = 'assembly-count';
        const wallText = `${pageMeshes.length} wall${pageMeshes.length === 1 ? '' : 's'}`;
        const roofText = roofMeshes.length ? `, ${roofMeshes.length} roof edge${roofMeshes.length === 1 ? '' : 's'}` : '';
        pageCount.textContent = wallText + roofText;

        pageRow.append(pageToggle, pageCheckbox, pageLabel, pageCount);
        children.appendChild(pageRow);

        // Walls list (collapsed by default)
        const pageChildren = document.createElement('div');
        pageChildren.className = 'assembly-page-children';
        pageChildren.dataset.pageChildren = page.id;

        pageMeshes.forEach((mesh, idx) => {
          const leaf = document.createElement('div');
          leaf.className = 'assembly-leaf';
          const entry = this.wallRegistry.get(mesh);
          leaf.textContent = entry?.label || `Wall ${idx + 1}`;
          pageChildren.appendChild(leaf);
        });

        if (roofMeshes.length) {
          const roofRow = document.createElement('div');
          roofRow.className = 'assembly-row';
          roofRow.style.paddingLeft = '0.35rem';

          const roofCheckbox = document.createElement('input');
          roofCheckbox.type = 'checkbox';
          roofCheckbox.dataset.pageId = page.id;
          roofCheckbox.dataset.sourceId = sourceId;
          roofCheckbox.dataset.roof = 'true';
          const roofVis = this.pageRoofVisibility.get(page.id);
          roofCheckbox.checked = roofVis !== false && roofMeshes.every((m) => m.visible !== false);

          const roofLabel = document.createElement('span');
          roofLabel.className = 'assembly-label';
          roofLabel.textContent = 'Roof';

          const roofCount = document.createElement('span');
          roofCount.className = 'assembly-count';
          roofCount.textContent = `${roofMeshes.length}`;

          roofRow.append(roofCheckbox, roofLabel, roofCount);
          pageChildren.appendChild(roofRow);
        }

        children.appendChild(pageChildren);
      });

      item.append(row, children);
      body.appendChild(item);
    });
  }

  private exportGlb() {
    const exporter = new GLTFExporter();
    exporter.parse(this.buildingGroup, (result) => {
      const arrayBuffer = result as ArrayBuffer;
      const blob = new Blob([arrayBuffer], { type: 'model/gltf-binary' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'house.glb';
      a.click();
      URL.revokeObjectURL(url);
    }, (err) => console.error(err), { binary: true });
  }

  private registerWallToSource(mesh: THREE.Mesh, sourceId: string, pageId: string, pageLabel: string) {
    const sourceVisible = this.sourceVisibility.get(sourceId);
    const pageVisible = this.pageVisibility.get(pageId);
    if (sourceVisible === false || pageVisible === false) mesh.visible = false;

    const list = this.sourceGroups.get(sourceId) || [];
    list.push(mesh);
    this.sourceGroups.set(sourceId, list);

    const pageEntry = this.pageGroups.get(pageId) || { id: pageId, meshes: [], roofMeshes: [], sourceId, label: pageLabel };
    pageEntry.meshes.push(mesh);
    this.pageGroups.set(pageId, pageEntry);
  }

  private registerRoofToSource(mesh: THREE.Mesh, sourceId: string, pageId: string, pageLabel: string) {
    const sourceVisible = this.sourceVisibility.get(sourceId);
    const pageVisible = this.pageRoofVisibility.get(pageId);
    if (sourceVisible === false || pageVisible === false) mesh.visible = false;

    const list = this.sourceGroups.get(sourceId) || [];
    list.push(mesh);
    this.sourceGroups.set(sourceId, list);

    const pageEntry = this.pageGroups.get(pageId) || { id: pageId, meshes: [], roofMeshes: [], sourceId, label: pageLabel };
    pageEntry.roofMeshes.push(mesh);
    this.pageGroups.set(pageId, pageEntry);
  }

  private setSourceVisibility(sourceId: string, visible: boolean) {
    this.sourceVisibility.set(sourceId, visible);
    const meshes = this.sourceGroups.get(sourceId) || [];
    meshes.forEach((mesh) => {
      const pageId = mesh.userData.pageId as string | undefined;
      const isRoof = mesh.userData.roof === true;
      const pageVisible = pageId
        ? (isRoof ? this.pageRoofVisibility.get(pageId) : this.pageVisibility.get(pageId))
        : undefined;
      const shouldShow = visible && (pageVisible !== false);
      mesh.visible = shouldShow;
    });
    this.renderAssemblyTree();
    this.addAutoFloorFromWalls();
    this.walkthroughController.syncEnvironment();
  }

  private setPageVisibility(pageId: string, visible: boolean) {
    this.pageVisibility.set(pageId, visible);
    const page = this.pageGroups.get(pageId);
    if (!page) return;
    const sourceVisible = this.sourceVisibility.get(page.sourceId);
    page.meshes.forEach((mesh) => {
      const shouldShow = (sourceVisible !== false) && visible;
      mesh.visible = shouldShow;
    });
    page.roofMeshes.forEach((mesh) => {
      const roofAllowed = this.pageRoofVisibility.get(pageId);
      const shouldShow = (sourceVisible !== false) && visible && (roofAllowed !== false);
      mesh.visible = shouldShow;
    });
    this.renderAssemblyTree();
    this.addAutoFloorFromWalls();
    this.walkthroughController.syncEnvironment();
  }

  private setPageRoofVisibility(pageId: string, visible: boolean) {
    this.pageRoofVisibility.set(pageId, visible);
    const page = this.pageGroups.get(pageId);
    if (!page) return;
    const sourceVisible = this.sourceVisibility.get(page.sourceId);
    page.roofMeshes.forEach((mesh) => {
      const shouldShow = (sourceVisible !== false) && visible;
      mesh.visible = shouldShow;
    });
    this.renderAssemblyTree();
    this.walkthroughController.syncEnvironment();
  }

  private initAssemblyPanel() {
    const collapseBtn = document.getElementById('assembly-collapse-btn') as HTMLButtonElement;
    const panel = document.getElementById('assembly-panel') as HTMLElement;
    const body = document.getElementById('assembly-body') as HTMLElement;

    collapseBtn.addEventListener('click', () => {
      this.assemblyCollapsed = !this.assemblyCollapsed;
      panel.classList.toggle('assembly-collapsed', this.assemblyCollapsed);
      collapseBtn.textContent = this.assemblyCollapsed ? 'Show' : 'Hide';
    });

    body.addEventListener('change', (e) => {
      const target = e.target as HTMLInputElement;
      if (target?.type === 'checkbox' && target.dataset.sourceId) {
        const visible = target.checked;
        if (target.dataset.pageId && target.dataset.roof === 'true') {
          this.setPageRoofVisibility(target.dataset.pageId, visible);
        } else if (target.dataset.pageId) {
          this.setPageVisibility(target.dataset.pageId, visible);
        } else {
          this.setSourceVisibility(target.dataset.sourceId, visible);
        }
      }
    });

    body.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest('.assembly-toggle') as HTMLElement | null;
      if (!btn) return;
      const sourceId = btn.dataset.sourceId;
      if (!sourceId) return;
      const child = body.querySelector(`[data-children="${sourceId}"]`) as HTMLElement | null;
      if (!child) return;
      const hidden = child.style.display === 'none';
      child.style.display = hidden ? 'flex' : 'none';
      btn.textContent = hidden ? '▾' : '▸';
    });

    body.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest('.assembly-page-toggle') as HTMLElement | null;
      if (!btn) return;
      const pageId = btn.dataset.pageId;
      if (!pageId) return;
      const child = body.querySelector(`[data-page-children="${pageId}"]`) as HTMLElement | null;
      if (!child) return;
      const hidden = child.style.display === 'none';
      child.style.display = hidden ? 'flex' : 'none';
      btn.textContent = hidden ? '▾' : '▸';
    });

    this.renderAssemblyTree();
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

  private normalizeTextureKey(rawKey: string | null | undefined): string {
    if (!rawKey) return 'none';
    return Object.prototype.hasOwnProperty.call(WALL_TEXTURE_URLS, rawKey) ? rawKey : 'none';
  }

  private getWallTextureKey(mesh: THREE.Mesh, entry: WallEntry): string {
    const settingsAny = entry.record.settings as Record<string, unknown>;
    const keyFromSettings = typeof settingsAny.texture === 'string' ? settingsAny.texture : null;
    const keyFromMesh = typeof mesh.userData.wallTextureKey === 'string' ? mesh.userData.wallTextureKey : null;
    return this.normalizeTextureKey(keyFromSettings || keyFromMesh);
  }

  private applyWallTexture(mesh: THREE.Mesh, entry: WallEntry, textureKey: string) {
    const normalizedKey = this.normalizeTextureKey(textureKey);
    const mat = mesh.material as THREE.MeshStandardMaterial;

    if (normalizedKey === 'none') {
      mat.map = null;
      mat.roughness = 0.8;
      mat.metalness = 0.1;
      mat.needsUpdate = true;
      mesh.userData.wallTextureKey = 'none';
      (entry.record.settings as any).texture = 'none';
      return;
    }

    let texture = this.wallTextureCache.get(normalizedKey);
    if (!texture) {
      texture = this.textureLoader.load(WALL_TEXTURE_URLS[normalizedKey]);
      texture.wrapS = THREE.RepeatWrapping;
      texture.wrapT = THREE.RepeatWrapping;
      texture.colorSpace = THREE.SRGBColorSpace;
      this.wallTextureCache.set(normalizedKey, texture);
    }

    texture.repeat.set(Math.max(1, entry.length / 6), Math.max(1, entry.height / 6));
    texture.needsUpdate = true;
    mat.map = texture;
    mat.color.set(0xffffff);
    mat.roughness = 0.55;
    mat.metalness = 0.04;
    mat.needsUpdate = true;
    mesh.userData.wallTextureKey = normalizedKey;
    (entry.record.settings as any).texture = normalizedKey;
  }

  private setModalTextureSelection(textureKey: string) {
    const normalized = this.normalizeTextureKey(textureKey);
    this.activeModalTextureKey = normalized;
    const swatches = document.querySelectorAll<HTMLButtonElement>('.texture-swatch');
    swatches.forEach((btn) => {
      const isActive = (btn.dataset.textureKey || 'none') === normalized;
      btn.classList.toggle('texture-swatch-active', isActive);
    });
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
      this.activeDisplayMesh.userData.walkthroughDoor = true;
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
    this.walkthroughController.syncEnvironment();
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
      this.loadDataFromJson(data, file.name);
    } catch (e) {
      fileInfo.textContent = 'Error: Could not parse JSON file.';
      console.error(e);
    }
  }

  private loadDataFromJson(data: any, fileName: string) {
    const fileInfo = document.querySelector('.file-info') as HTMLElement;
    const looksLikePlanPages = Array.isArray(data.pages) && data.pages.some((p: any) => {
      const e = p?.entities;
      return e && ((e.walls && e.walls.length) || e.windows_and_doors_floor_plans || e.roofing);
    });

    if (looksLikePlanPages) {
      this.renderPlanPages(data as PlanJson, fileName);
    } else if (data.records && Array.isArray(data.records)) {
      this.renderNewFormat(data as NewJsonData, fileName);
    } else if (data.walls && Array.isArray(data.walls)) {
      if (data.walls.length > 0 && data.walls[0].start && data.walls[0].end) {
        this.renderLineFormat(data, fileName);
      } else {
        this.renderOldFormat(data as OldJsonData, fileName);
      }
    } else {
      fileInfo.textContent = 'Error: Unrecognised JSON format.';
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // PLAN PAGE renderer (pages[].entities.walls + windows_and_doors_floor_plans)
  // ═══════════════════════════════════════════════════════════════════════════
  private renderPlanPages(data: PlanJson, fileName: string) {
    this.clearScene();

    const fileInfo = document.querySelector('.file-info') as HTMLElement;
    const pages = (data.pages || []).filter((p) => p.entities && ((p.entities.walls && p.entities.walls.length) || p.entities?.windows_and_doors_floor_plans || p.entities?.roofing));
    if (!pages.length) {
      fileInfo.textContent = 'No entities found in plan file.';
      return;
    }

    type WallDatum = {
      x1: number; y1: number; x2: number; y2: number;
      thicknessIn: number; heightFt: number; label: string; sourceId: string; pageId: string; pageLabel: string;
    };

    const walls: WallDatum[] = [];
    const openings: OpeningBox[] = [];
    const roofEdgeQueue: Array<{ edges: any; sourceId: string; pageId: string; pageLabel: string }> = [];
    const roofObjQueue: Array<{ roof: any; sourceId: string; pageId: string; pageLabel: string }> = [];
    let roofEdgeCount = 0;
    let roofPlaneCount = 0;

    pages.forEach((page) => {
      const entities = page.entities || {};
      const sourceId = page.source_file_id || 'Unknown Source';
      const pageLabel = `Page ${page.page_number_in_source_file ?? page.page_index ?? ''}`.trim();
      const pageId = `${sourceId}::${page.page_number_in_source_file ?? page.page_index ?? Math.random().toString(36).slice(2)}`;
      (entities.walls || []).forEach((wall, idx) => {
        const coords = wall.geometry?.coordinates;
        if (!coords) return;
        const x1 = Number((coords as any).x1);
        const y1 = Number((coords as any).y1);
        const x2 = Number((coords as any).x2);
        const y2 = Number((coords as any).y2);
        if (!Number.isFinite(x1) || !Number.isFinite(y1) || !Number.isFinite(x2) || !Number.isFinite(y2)) return;

        const thicknessIn = Number(wall.properties?.thickness_inches) || 6;
        const heightFt = parseFeetInches(wall.properties?.wall_height, 9);
        const label = normalizePlanLabel(wall.category || wall.properties?.category || wall.properties?.floor_label);

        walls.push({
          x1,
          y1,
          x2,
          y2,
          thicknessIn,
          heightFt,
          label: label || `Wall ${idx + 1}`,
          sourceId,
          pageId,
          pageLabel,
        });
      });

      const boxes = this.extractOpeningBoxes(entities.windows_and_doors_floor_plans);
      openings.push(...boxes);

      if (entities.roofing?.EdgesOnly) {
        roofEdgeQueue.push({ edges: entities.roofing.EdgesOnly, sourceId, pageId, pageLabel });
        roofEdgeCount += Array.isArray(entities.roofing.EdgesOnly.keypoints) ? entities.roofing.EdgesOnly.keypoints.length : 0;
      }
      if (entities.roofing) {
        roofObjQueue.push({ roof: entities.roofing, sourceId, pageId, pageLabel });
        roofPlaneCount += Array.isArray(entities.roofing.primary_roof?.roof_planes) ? entities.roofing.primary_roof.roof_planes.length : 0;
        roofPlaneCount += Array.isArray(entities.roofing.cross_roof?.roof_planes) ? entities.roofing.cross_roof.roof_planes.length : 0;
      }
    });

    if (!walls.length && !roofEdgeCount) {
      fileInfo.textContent = 'No walls or roof edges present in plan file.';
      return;
    }

    // Compute bounds for centering
    let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
    const extendBounds = (x: number, y: number) => {
      minX = Math.min(minX, x);
      maxX = Math.max(maxX, x);
      minY = Math.min(minY, y);
      maxY = Math.max(maxY, y);
    };

    walls.forEach((w) => {
      extendBounds(w.x1, w.y1);
      extendBounds(w.x2, w.y2);
    });

    roofEdgeQueue.forEach(({ edges }) => {
      const kp: any[] = edges?.keypoints || [];
      kp.forEach((pair) => {
        if (!Array.isArray(pair) || pair.length < 2) return;
        const x1 = Number(pair[0]?.[0]);
        const y1 = Number(pair[0]?.[1]);
        const x2 = Number(pair[1]?.[0]);
        const y2 = Number(pair[1]?.[1]);
        if ([x1, y1, x2, y2].some((v) => !Number.isFinite(v))) return;
        extendBounds(x1, y1);
        extendBounds(x2, y2);
      });
    });

    roofObjQueue.forEach(({ roof }) => {
      const takePt = (p: any) => {
        if (!p) return;
        const x = Number(p.x ?? p[0]);
        const y = Number(p.y ?? p[1]);
        if (Number.isFinite(x) && Number.isFinite(y)) extendBounds(x, y);
      };
      const collectBoundary = (boundary: any[]) => {
        if (!Array.isArray(boundary)) return;
        boundary.forEach(takePt);
      };

      const primary = roof.primary_roof;
      const cross = roof.cross_roof;
      [primary, cross].forEach((group) => {
        if (!group) return;
        if (Array.isArray(group.roof_planes)) group.roof_planes.forEach((pl: any) => collectBoundary(pl?.boundary));
        if (Array.isArray(group.eaves)) group.eaves.forEach((ev: any) => {
          takePt(ev?.start ?? ev?.[0]);
          takePt(ev?.end ?? ev?.[1]);
        });
        if (group.ridge_line) {
          if (Array.isArray(group.ridge_line)) group.ridge_line.forEach(takePt);
          else { takePt(group.ridge_line.start); takePt(group.ridge_line.end); }
        }
        if (Array.isArray(group.valleys)) group.valleys.forEach((vl: any) => { takePt(vl?.start ?? vl?.[0]); takePt(vl?.end ?? vl?.[1]); });
        if (Array.isArray(group.gable_ends)) group.gable_ends.forEach((g: any) => {
          if (g?.wall_line) {
            takePt({ x: g.wall_line.x1, y: g.wall_line.y1 });
            takePt({ x: g.wall_line.x2, y: g.wall_line.y2 });
          }
          takePt(g?.peak);
        });
      });
    });
    openings.forEach((b) => {
      if (!Number.isFinite(b.xmin) || !Number.isFinite(b.xmax) || !Number.isFinite(b.ymin) || !Number.isFinite(b.ymax)) return;
      minX = Math.min(minX, b.xmin, b.xmax);
      maxX = Math.max(maxX, b.xmin, b.xmax);
      minY = Math.min(minY, b.ymin, b.ymax);
      maxY = Math.max(maxY, b.ymin, b.ymax);
    });

    const cx = (minX + maxX) / 2;
    const cy = (minY + maxY) / 2;

    if (!Number.isFinite(cx) || !Number.isFinite(cy)) {
      fileInfo.textContent = 'Invalid geometry in plan file (non-finite coordinates).';
      return;
    }

    const wallTypeCounts: Record<string, number> = {};
    const typeCounts: Record<string, number> = { wall: walls.length };
    const totalRoof = roofEdgeCount + roofPlaneCount;
    if (totalRoof) typeCounts['roof'] = totalRoof;

    type WallSeg = {
      mesh: THREE.Mesh;
      a: THREE.Vector3;
      b: THREE.Vector3;
      thickness: number;
      height: number;
      label: string;
      rotY: number;
    };

    const wallSegments: WallSeg[] = [];

    walls.forEach((w, idx) => {
      const a = new THREE.Vector3((w.x1 - cx) * SCALE, 0, (w.y1 - cy) * SCALE);
      const b = new THREE.Vector3((w.x2 - cx) * SCALE, 0, (w.y2 - cy) * SCALE);

      const dir = new THREE.Vector3().subVectors(b, a);
      const length = dir.length();
      if (!Number.isFinite(length) || length < 0.001) return;

      const thickness = (w.thicknessIn || 6) * SCALE;
      const height = w.heightFt || 9;
      if (!Number.isFinite(thickness) || !Number.isFinite(height)) return;

      const geo = new THREE.BoxGeometry(length, height, thickness);
      const mat = new THREE.MeshStandardMaterial({ color: 0xd9d9d9, metalness: 0.05, roughness: 0.85 });
      const mesh = new THREE.Mesh(geo, mat);
      mesh.userData.sourceId = w.sourceId;
      mesh.userData.pageId = w.pageId;

      const mid = new THREE.Vector3().copy(a).lerp(b, 0.5);
      mid.y = height / 2;
      const rotY = -Math.atan2(dir.z, dir.x);

      mesh.position.copy(mid);
      mesh.rotation.y = rotY;
      mesh.castShadow = true;
      mesh.receiveShadow = true;
      mesh.updateMatrix();

      this.buildingGroup.add(mesh);

      const label = w.label || `Wall ${idx + 1}`;
      wallTypeCounts[label] = (wallTypeCounts[label] || 0) + 1;

      const fakeRecord: NewRecord = {
        materialType: 'wall',
        settings: { name: label, type: label, height: String(height), floor_level: 'default' },
        coordinates_real_world: [[w.x1, w.y1], [w.x2, w.y2]],
        scale_factor_float: SCALE,
      } as NewRecord;

      this.wallRegistry.set(mesh, {
        pts: [[a.x, a.z], [b.x, b.z]],
        cx, cy,
        height,
        thickness,
        length,
        baseElev: 0,
        label,
        wallType: label,
        record: fakeRecord,
        originalColor: 0xd9d9d9,
        worldPos: mid.clone(),
        worldRotY: rotY,
      });

      wallSegments.push({ mesh, a, b, thickness, height, label, rotY });
      this.registerWallToSource(mesh, w.sourceId, w.pageId, w.pageLabel);
    });

    // Render roof edges (if present) for each page/source
    roofEdgeQueue.forEach((task) => {
      this.renderRoofEdges(task.edges, cx, cy, task.sourceId, task.pageId, task.pageLabel);
      this.renderRoofing(task.edges, cx, cy, task.sourceId, task.pageId, task.pageLabel);
    });
    roofObjQueue.forEach((task) => {
      this.renderRoofing(task.roof, cx, cy, task.sourceId, task.pageId, task.pageLabel);
    });

    const openingHeightDefault = 7; // feet

    if (!wallSegments.length) {
      this.updateStatsPanel(typeCounts, wallTypeCounts);
      this.addAutoFloorFromWalls();
      fileInfo.textContent = `Plan ${fileName} — ${walls.length} walls, ${openings.length} openings`;
      this.updateLegendForOldFormat();
      this.frameCamera();
      this.walkthroughController.syncEnvironment();
      return;
    }

    openings.forEach((box) => {
      const centerX = (box.xmin + box.xmax) / 2;
      const centerY = (box.ymin + box.ymax) / 2;
      const width = Math.max(0.1, (box.xmax - box.xmin) * SCALE);

      const center = new THREE.Vector3((centerX - cx) * SCALE, 0, (centerY - cy) * SCALE);

      if (!Number.isFinite(center.x) || !Number.isFinite(center.z) || !Number.isFinite(width)) return;

      let nearest: WallSeg | null = null;
      let nearestDist = Infinity;

      wallSegments.forEach((seg) => {
        const ab = new THREE.Vector3().subVectors(seg.b, seg.a);
        const ap = new THREE.Vector3().subVectors(center, seg.a);
        const t = Math.max(0, Math.min(1, ap.dot(ab) / Math.max(ab.lengthSq(), 1e-6)));
        const proj = seg.a.clone().addScaledVector(ab, t);
        const dist = proj.distanceTo(center);
        if (dist < nearestDist) {
          nearestDist = dist;
          nearest = seg;
        }
      });

      if (!nearest) return;

      const nearestWall = nearest as WallSeg;

      const openingHeight = openingHeightDefault;
      const thickness = nearestWall.thickness + 0.05; // ensure cut fully passes through

      const geom = new THREE.BoxGeometry(width, openingHeight, thickness);
      const mat = new THREE.MeshStandardMaterial({ color: 0x93c5fd, transparent: true, opacity: 0.35 });
      const openingMesh = new THREE.Mesh(geom, mat);
      openingMesh.position.set(center.x, openingHeight / 2, center.z);
      openingMesh.rotation.y = nearestWall.rotY;
      openingMesh.updateMatrix();

      // Subtract from nearest wall using CSG while keeping the same mesh reference
      const wallMesh = nearestWall.mesh;
      const wallCSG = CSG.fromMesh(wallMesh);
      const openingCSG = CSG.fromMesh(openingMesh);
      const result = wallCSG.subtract(openingCSG);
      const resultMesh = CSG.toMesh(result, wallMesh.matrix, wallMesh.material as THREE.Material);

      wallMesh.geometry.dispose();
      wallMesh.geometry = resultMesh.geometry;
      wallMesh.updateMatrix();
    });

    this.updateStatsPanel(typeCounts, wallTypeCounts);
    this.addAutoFloorFromWalls();
    this.renderAssemblyTree();

    fileInfo.textContent = `Plan ${fileName} — ${walls.length} walls, ${openings.length} openings`;
    this.updateLegendForOldFormat();
    this.frameCamera();
    this.walkthroughController.syncEnvironment();
  }

  private extractOpeningBoxes(floorPlan: any): OpeningBox[] {
    if (!floorPlan) return [];
    const boxes: OpeningBox[] = [];

    const candidates = floorPlan.boxes || floorPlan.HeadersOnly?.boxes || floorPlan.Boxes || [];
    if (Array.isArray(candidates)) {
      candidates.forEach((box: any) => {
        const parsed = toOpeningBox(box);
        if (parsed) boxes.push(parsed);
      });
    }
    return boxes;
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
    this.walkthroughController.syncEnvironment();
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
    const entry = this.wallRegistry.get(mesh)!;
    this.applyWallTexture(mesh, entry, typeof settings.texture === 'string' ? settings.texture : 'none');

    return mesh;
  }

  // ─── Roof (plan-format, boundary-defined) ───────────────────────────────
  private renderRoofing(roofing: any, cx: number, cy: number, sourceId: string, pageId: string, pageLabel: string) {
    if (!roofing) return;

    const global = roofing.global_properties || {};
    const eaveHeight = Number(global.eave_height ?? 9);
    const ridgeHeight = Number(global.ridge_height ?? eaveHeight + 4);
    // defaults available if we need to procedurally generate planes
    const defaultSlopeDeg = Number(global.default_slope_degrees ?? 30);
    const roofThickness = Number(global.thickness ?? 0.5);
    void defaultSlopeDeg; void roofThickness;

    const material = new THREE.MeshStandardMaterial({ color: 0xeb5f0c, metalness: 0.15, roughness: 0.65, side: THREE.DoubleSide, transparent: true, opacity: 0.9 });
    const lineMat = new THREE.LineBasicMaterial({ color: 0xf97316, linewidth: 2, transparent: true, opacity: 0.95 });

    const toV3 = (p: any): THREE.Vector3 | null => {
      if (!p) return null;
      const x = Number(Array.isArray(p) ? p[0] : p.x ?? p.x1 ?? p.start?.x ?? p.end?.x);
      const y = Number(Array.isArray(p) ? p[1] : p.y ?? p.y1 ?? p.start?.y ?? p.end?.y);
      const z = Number(Array.isArray(p) ? p[2] : p.z ?? p.z1 ?? p.start?.z ?? p.end?.z);
      if (![x, y, z].every(Number.isFinite)) return null;
      return new THREE.Vector3((x - cx) * SCALE, z * SCALE, (y - cy) * SCALE);
    };

    const addLine = (a: THREE.Vector3, b: THREE.Vector3) => {
      const geom = new THREE.BufferGeometry().setFromPoints([a, b]);
      const line = new THREE.Line(geom, lineMat.clone());
      line.userData = { roof: true, sourceId, pageId };
      this.buildingGroup.add(line);
      this.registerRoofToSource(line as unknown as THREE.Mesh, sourceId, pageId, pageLabel);
    };

    const buildPlane = (boundary: any[]) => {
      if (!Array.isArray(boundary) || boundary.length < 3) return;
      const pts: THREE.Vector3[] = [];
      boundary.forEach((pt) => {
        const v = toV3(pt);
        if (v) pts.push(v);
      });
      if (pts.length < 3) return;

      // Simple fan triangulation
      const positions: number[] = [];
      for (let i = 1; i < pts.length - 1; i++) {
        const tri = [pts[0], pts[i], pts[i + 1]];
        tri.forEach((v) => positions.push(v.x, v.y, v.z));
      }

      const geom = new THREE.BufferGeometry();
      geom.setAttribute('position', new THREE.Float32BufferAttribute(positions, 3));
      geom.computeVertexNormals();

      const mesh = new THREE.Mesh(geom, material.clone());
      mesh.userData = { roof: true, sourceId, pageId };
      mesh.castShadow = true;
      mesh.receiveShadow = true;
      this.buildingGroup.add(mesh);
      this.registerRoofToSource(mesh, sourceId, pageId, pageLabel);
    };

    const handleRoofGroup = (roofObj: any) => {
      if (!roofObj) return;

      // Ridge visualization
      const ridge = roofObj.ridge_line;
      if (Array.isArray(ridge) && ridge.length >= 2) {
        for (let i = 0; i < ridge.length - 1; i++) {
          const a = toV3(ridge[i]);
          const b = toV3(ridge[i + 1]);
          if (a && b) addLine(a, b);
        }
      } else if (ridge?.start && ridge?.end) {
        const a = toV3(ridge.start);
        const b = toV3(ridge.end);
        if (a && b) addLine(a, b);
      }

      // Eaves lines at eave height if provided
      const eaves = roofObj.eaves;
      if (Array.isArray(eaves)) {
        eaves.forEach((pair: any) => {
          const start = pair?.start ?? pair?.[0];
          const end = pair?.end ?? pair?.[1];
          if (!start || !end) return;
          const a = toV3({ x: start.x ?? start[0], y: start.y ?? start[1], z: eaveHeight });
          const b = toV3({ x: end.x ?? end[0], y: end.y ?? end[1], z: eaveHeight });
          if (a && b) addLine(a, b);
        });
      }

      // Roof planes
      const planes = roofObj.roof_planes || [];
      planes.forEach((pl: any) => {
        if (Array.isArray(pl?.boundary)) buildPlane(pl.boundary);
      });

      // Gables
      const gables = roofObj.gable_ends || [];
      gables.forEach((g: any) => {
        if (!g?.wall_line || !g?.peak) return;
        const wl = g.wall_line;
        const a = wl?.x1 !== undefined ? toV3({ x: wl.x1, y: wl.y1, z: eaveHeight }) : toV3(wl?.[0]);
        const b = wl?.x2 !== undefined ? toV3({ x: wl.x2, y: wl.y2, z: eaveHeight }) : toV3(wl?.[1]);
        const c = toV3(g.peak);
        if (!a || !b || !c) return;
        const geom = new THREE.BufferGeometry();
        geom.setFromPoints([a, b, c]);
        geom.setIndex([0, 1, 2]);
        geom.computeVertexNormals();
        const mesh = new THREE.Mesh(geom, material.clone());
        mesh.userData = { roof: true, sourceId, pageId };
        this.buildingGroup.add(mesh);
        this.registerRoofToSource(mesh, sourceId, pageId, pageLabel);
      });

      // Valleys: visualize lines
      const valleys = roofObj.valleys || roofObj.valley_lines || [];
      valleys.forEach((pair: any) => {
        const start = pair?.start ?? pair?.[0];
        const end = pair?.end ?? pair?.[1];
        const a = toV3(start);
        const b = toV3(end);
        if (a && b) addLine(a, b);
      });
    };

    // Primary and cross roofs
    handleRoofGroup(roofing.primary_roof);
    handleRoofGroup(roofing.cross_roof);

    // If only global props and no planes, fallback: draw ridge/eave box using defaults
    if (!roofing.primary_roof && roofing.global_properties) {
      const size = 10 * SCALE;
      const a = new THREE.Vector3(-size, eaveHeight, -size);
      const b = new THREE.Vector3(size, eaveHeight, -size);
      const c = new THREE.Vector3(0, ridgeHeight, 0);
      const geom = new THREE.BufferGeometry();
      geom.setFromPoints([a, b, c]);
      geom.setIndex([0, 1, 2]);
      geom.computeVertexNormals();
      const mesh = new THREE.Mesh(geom, material.clone());
      mesh.userData = { roof: true, sourceId, pageId };
      this.buildingGroup.add(mesh);
      this.registerRoofToSource(mesh, sourceId, pageId, pageLabel);
    }
  }

  private renderRoofEdges(edges: any, cx: number, cy: number, sourceId: string, pageId: string, pageLabel: string) {
    if (!edges || !Array.isArray(edges.keypoints)) return;

    const keypoints: any[] = edges.keypoints || [];
    const pitchs: any[] = Array.isArray(edges.pitchs) ? edges.pitchs : [];
    const classCandidates: any[] = Array.isArray(edges.class_ids) ? edges.class_ids : [];
    const classList: any[] = classCandidates.find((arr) => Array.isArray(arr) && arr.length === keypoints.length) || [];

    const material = new THREE.LineBasicMaterial({ color: 0xf97316, linewidth: 2, transparent: true, opacity: 0.9 });
    const baseHeight = 9; // feet

    keypoints.forEach((pair, idx) => {
      if (!Array.isArray(pair) || pair.length < 2) return;
      const p1 = pair[0];
      const p2 = pair[1];
      const x1 = Number(p1?.[0]);
      const y1 = Number(p1?.[1]);
      const x2 = Number(p2?.[0]);
      const y2 = Number(p2?.[1]);
      if (!Number.isFinite(x1) || !Number.isFinite(y1) || !Number.isFinite(x2) || !Number.isFinite(y2)) return;

      const v1 = new THREE.Vector3((x1 - cx) * SCALE, baseHeight, (y1 - cy) * SCALE);
      const v2 = new THREE.Vector3((x2 - cx) * SCALE, baseHeight, (y2 - cy) * SCALE);

      const run = new THREE.Vector2(v2.x - v1.x, v2.z - v1.z).length();
      const pitchCandidate = pitchs[idx];
      let pitch = 9;
      if (Array.isArray(pitchCandidate) && pitchCandidate.length) pitch = parseFloat(pitchCandidate[0]) || 9;
      else if (typeof pitchCandidate === 'string' || typeof pitchCandidate === 'number') pitch = parseFloat(pitchCandidate as any) || 9;

      const deltaH = run * (pitch / 12);
      const classId = Array.isArray(classList) ? classList[idx] : 0;
      const isHorizontal = classId === 1 || classId === 4;

      if (!isHorizontal) v2.y = baseHeight + deltaH;

      const geom = new THREE.BufferGeometry().setFromPoints([v1, v2]);
      const line = new THREE.Line(geom, material.clone());
      line.userData = { roof: true, sourceId, pageId };
      this.buildingGroup.add(line);
      this.registerRoofToSource(line as unknown as THREE.Mesh, sourceId, pageId, pageLabel);
    });
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
      } else {
        const doorPortalMesh = mesh.clone();
        doorPortalMesh.geometry = new THREE.BoxGeometry(cutoutWidth, cutoutHeight, 0.45);
        doorPortalMesh.visible = false;
        doorPortalMesh.userData.walkthroughDoor = true;
        this.buildingGroup.add(doorPortalMesh);
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
        settings: { id: w.id, name: wallTypeLabel, type: w.room, height: String(wallHeight), floor_level: 'default' },
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
    this.walkthroughController.syncEnvironment();
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
      const mat = new THREE.MeshStandardMaterial({ color: floorColor, side: THREE.DoubleSide, roughness: 0.85, metalness: 0.02 });
      const mesh = new THREE.Mesh(geo, mat);
      // Box geometry is centered — place it so top of floor sits at defaultFloorHeight
      mesh.position.set(((x1 + x2) / 2 - cx) * SCALE, defaultFloorHeight + (floorThicknessFt / 2), ((y1 + y2) / 2 - cy) * SCALE);
      mesh.userData.walkthroughFloor = true;
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
      mesh.userData.walkthroughDoor = true;
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
    this.walkthroughController.syncEnvironment();
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
    return Array.from(this.wallRegistry.keys()).filter((mesh) => mesh.visible !== false);
  }

  private getFloorMeshes(): THREE.Mesh[] {
    return this.buildingGroup.children
      .filter((child) => (child as THREE.Mesh).isMesh)
      .map((child) => child as THREE.Mesh)
      .filter((mesh) => mesh.userData && mesh.userData.walkthroughFloor === true);
  }

  private getDoorOpeningMeshes(): THREE.Mesh[] {
    return this.buildingGroup.children
      .filter((child) => (child as THREE.Mesh).isMesh)
      .map((child) => child as THREE.Mesh)
      .filter((mesh) => mesh.userData && mesh.userData.walkthroughDoor === true);
  }

  private updateGuidedNavUI(state: GuidedState) {
    const wrap = document.getElementById('guided-nav-controls') as HTMLElement;
    const label = document.getElementById('guided-stop-label') as HTMLElement;
    const prev = document.getElementById('guided-prev-btn') as HTMLButtonElement;
    const next = document.getElementById('guided-next-btn') as HTMLButtonElement;

    wrap.style.display = state.visible ? 'flex' : 'none';
    if (!state.visible) return;

    label.textContent = `Stop ${state.current} / ${state.total}`;
    prev.disabled = !state.canPrev;
    next.disabled = !state.canNext;
    prev.style.opacity = state.canPrev ? '1' : '0.55';
    next.style.opacity = state.canNext ? '1' : '0.55';
    prev.style.cursor = state.canPrev ? 'pointer' : 'not-allowed';
    next.style.cursor = state.canNext ? 'pointer' : 'not-allowed';
  }

  private onCanvasClick(e: PointerEvent) {
    if (e.button !== 0) return; // left click only
    if (this.walkthroughController.handleGuidedPointerDown(e)) {
      e.preventDefault();
      return;
    }

    const canvas = this.renderer.domElement;
    const rect = canvas.getBoundingClientRect();
    this.pointer.x = ((e.clientX - rect.left) / rect.width) * 2 - 1;
    this.pointer.y = -((e.clientY - rect.top) / rect.height) * 2 + 1;

    this.raycaster.setFromCamera(this.pointer, this.camera);

    if (this.walkthroughController.tryAddStopFromRay(this.raycaster)) {
      e.stopPropagation();
      return;
    }

    const wallMeshes = this.getWallMeshes();
    const hits = this.raycaster.intersectObjects(wallMeshes, false);

    if (hits.length > 0) {
      const hit = hits[0].object as THREE.Mesh;
      if (this.placementMode !== 'none') {
        this.startCutoutPlacement(hit);
      } else if (this.wallEditEnabled) {
        this.selectWall(hit);
      } else {
        this.deselectWall();
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
    const wallHits = this.raycaster.intersectObjects(this.getWallMeshes(), false);
    const floorHits = this.raycaster.intersectObjects(this.getFloorMeshes(), false);
    const shouldShowHover = this.placementMode !== 'none' || this.wallEditEnabled;
    const showForPathEdit = floorHits.length > 0 && (document.getElementById('guided-path-edit-toggle') as HTMLInputElement).checked;
    document.body.classList.toggle('wall-hover', (shouldShowHover && wallHits.length > 0) || showForPathEdit);
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
    const mat = mesh.material as THREE.MeshStandardMaterial;

    labelEl.textContent = entry.label.replace(/(^\w+\s+){2}/, '') || entry.label;
    const settingsAny = entry.record.settings as Record<string, unknown>;
    const wallId = settingsAny.id ?? settingsAny.wall_id ?? settingsAny.room_id ?? null;
    typeEl.textContent = wallId ? `ID: ${String(wallId)}` : entry.wallType;
    heightIn.value = entry.height.toFixed(2);
    lengthIn.value = entry.length.toFixed(2);
    widthIn.value = entry.thickness.toFixed(2);
    this.setModalColorControls(`#${mat.color.getHexString()}`);
    this.setModalTextureSelection(this.getWallTextureKey(mesh, entry));

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
    const textureKey = this.normalizeTextureKey(this.activeModalTextureKey);
    const beforeSnapshot = this.captureWallSnapshot(mesh, entry);

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
    this.applyWallTexture(mesh, entry, textureKey);
    // Update worldPos Y to match the new height centre
    entry.worldPos.y = mesh.position.y;
    this.addAutoFloorFromWalls();
    this.refreshEstimateIfOpen();
    this.walkthroughController.syncEnvironment();

    const afterSnapshot = this.captureWallSnapshot(mesh, entry);
    this.pushHistory(mesh, beforeSnapshot, afterSnapshot);

    // Close modal and deselect
    this.closeModal();
  }

  private captureWallSnapshot(mesh: THREE.Mesh, entry: WallEntry): WallEditSnapshot {
    const mat = mesh.material as THREE.MeshStandardMaterial;
    return {
      height: entry.height,
      length: entry.length,
      thickness: entry.thickness,
      colorHex: `#${mat.color.getHexString()}`,
      textureKey: this.getWallTextureKey(mesh, entry),
      baseElev: entry.baseElev,
      worldPos: entry.worldPos.clone(),
      worldRotY: entry.worldRotY,
    };
  }

  private applyWallSnapshot(mesh: THREE.Mesh, snapshot: WallEditSnapshot) {
    const entry = this.wallRegistry.get(mesh);
    if (!entry) return;

    mesh.geometry.dispose();
    mesh.geometry = new THREE.BoxGeometry(snapshot.length, snapshot.height, snapshot.thickness);
    mesh.position.copy(snapshot.worldPos);
    mesh.position.y = snapshot.baseElev + snapshot.height / 2;
    mesh.rotation.y = snapshot.worldRotY;

    const mat = mesh.material as THREE.MeshStandardMaterial;
    mat.color.set(snapshot.colorHex);

    entry.height = snapshot.height;
    entry.length = snapshot.length;
    entry.thickness = snapshot.thickness;
    entry.baseElev = snapshot.baseElev;
    entry.worldPos.copy(snapshot.worldPos);
    entry.worldPos.y = mesh.position.y;
    entry.worldRotY = snapshot.worldRotY;
    entry.originalColor = mat.color.getHex();
    (entry.record.settings as any).color = snapshot.colorHex;
    this.applyWallTexture(mesh, entry, snapshot.textureKey);
  }

  private pushHistory(mesh: THREE.Mesh, before: WallEditSnapshot, after: WallEditSnapshot) {
    const isSame =
      before.height === after.height &&
      before.length === after.length &&
      before.thickness === after.thickness &&
      before.colorHex === after.colorHex &&
      before.textureKey === after.textureKey;
    if (isSame) return;

    this.undoStack.push({ mesh, before, after });
    if (this.undoStack.length > this.maxHistorySize) this.undoStack.shift();
    this.redoStack = [];
  }

  private undoWallEdit() {
    const action = this.undoStack.pop();
    if (!action) return;
    if (!this.wallRegistry.has(action.mesh)) return;

    this.applyWallSnapshot(action.mesh, action.before);
    this.redoStack.push(action);
    if (this.redoStack.length > this.maxHistorySize) this.redoStack.shift();
    this.addAutoFloorFromWalls();
    this.refreshEstimateIfOpen();
    this.walkthroughController.syncEnvironment();
  }

  private redoWallEdit() {
    const action = this.redoStack.pop();
    if (!action) return;
    if (!this.wallRegistry.has(action.mesh)) return;

    this.applyWallSnapshot(action.mesh, action.after);
    this.undoStack.push(action);
    if (this.undoStack.length > this.maxHistorySize) this.undoStack.shift();
    this.addAutoFloorFromWalls();
    this.refreshEstimateIfOpen();
    this.walkthroughController.syncEnvironment();
  }

  private handleUndoRedoShortcuts(e: KeyboardEvent) {
    const target = e.target as HTMLElement | null;
    if (target) {
      const tag = target.tagName;
      const isEditable = target.isContentEditable || tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT';
      if (isEditable) return;
    }

    const isMac = /Mac|iPhone|iPad|iPod/.test(navigator.platform);
    const primary = isMac ? e.metaKey : e.ctrlKey;
    if (!primary || e.altKey) return;
    if (e.key.toLowerCase() !== 'z' && e.key.toLowerCase() !== 'y') return;

    if (e.key.toLowerCase() === 'z') {
      e.preventDefault();
      if (isMac && e.shiftKey) this.redoWallEdit();
      else if (!isMac && e.shiftKey) this.redoWallEdit();
      else this.undoWallEdit();
      return;
    }

    if (!isMac && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      this.redoWallEdit();
    }
  }

  // ─── Scene helpers ────────────────────────────────────────────────────────
  private clearScene() {
    this.wallRegistry.clear();
    this.selectedWall = null;
    this.undoStack = [];
    this.redoStack = [];
    this.sourceGroups.clear();
    this.pageGroups.clear();
    this.sourceVisibility.clear();
    this.pageVisibility.clear();
    this.pageRoofVisibility.clear();
    this.walkthroughController.stop();
    this.walkthroughController.clearCustomStops();
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

    this.renderAssemblyTree();
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
    });
    const floorMesh = new THREE.Mesh(floorGeo, floorMat);
    floorMesh.name = 'auto-generated-floor';
    floorMesh.userData.walkthroughFloor = true;
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
    const delta = this.clock.getDelta();
    if (!this.walkthroughController.isActive()) this.controls.update();
    this.walkthroughController.update(delta);
    this.renderer.render(this.scene, this.camera);
  }
}

new HouseViewer();
