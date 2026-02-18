import './style.css';
import * as THREE from 'three';
import { OrbitControls } from 'three/examples/jsm/controls/OrbitControls.js';
import { inject } from '@vercel/analytics';

// Initialize Vercel Analytics
inject();

interface WallData {
  class: string;
  confidence: number;
  bbox: {
    x1: number;
    y1: number;
    x2: number;
    y2: number;
  };
  center: {
    x: number;
    y: number;
  };
  area: number;
}

interface JsonData {
  walls: WallData[];
}

class HouseViewer {
  private scene: THREE.Scene;
  private camera: THREE.PerspectiveCamera;
  private renderer: THREE.WebGLRenderer;
  private controls: OrbitControls;
  private wallsGroup: THREE.Group;
  private wallHeight: number = 2.5;
  private wallMaterials!: Record<string, THREE.Material>;

  constructor() {
    this.scene = new THREE.Scene();
    this.scene.background = new THREE.Color(0x0f172a);

    const canvas = document.querySelector('#three-canvas') as HTMLCanvasElement;
    this.renderer = new THREE.WebGLRenderer({
      canvas,
      antialias: true,
    });
    this.renderer.setSize(window.innerWidth, window.innerHeight);
    this.renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));

    this.camera = new THREE.PerspectiveCamera(75, window.innerWidth / window.innerHeight, 0.1, 1000);
    this.camera.position.set(20, 20, 20);

    this.controls = new OrbitControls(this.camera, this.renderer.domElement);
    this.controls.enableDamping = true;

    this.wallsGroup = new THREE.Group();
    this.scene.add(this.wallsGroup);

    this.initLights();
    this.initMaterials();
    this.initEventListeners();
    this.animate();
  }

  private initLights() {
    const ambientLight = new THREE.AmbientLight(0xffffff, 0.6);
    this.scene.add(ambientLight);

    const directionalLight = new THREE.DirectionalLight(0xffffff, 0.8);
    directionalLight.position.set(100, 100, 50);
    this.scene.add(directionalLight);

    const pointLight = new THREE.PointLight(0x818cf8, 0.5);
    pointLight.position.set(-50, 50, -50);
    this.scene.add(pointLight);
  }

  private initMaterials() {
    this.wallMaterials = {
      perimeter_wall: new THREE.MeshStandardMaterial({ color: 0xef4444, metalness: 0.1, roughness: 0.8 }),
      interior_wall: new THREE.MeshStandardMaterial({ color: 0x3b82f6, metalness: 0.1, roughness: 0.8 }),
      foundation_wall: new THREE.MeshStandardMaterial({ color: 0x64748b, metalness: 0.1, roughness: 0.8 }),
      knee_wall: new THREE.MeshStandardMaterial({ color: 0xf59e0b, metalness: 0.1, roughness: 0.8 }),
      default: new THREE.MeshStandardMaterial({ color: 0xffffff, metalness: 0.1, roughness: 0.8 }),
    };
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
    fileInfo.textContent = `Loading: ${file.name}...`;

    try {
      const text = await file.text();
      let data: JsonData;
      try {
        data = JSON.parse(text);
      } catch (e) {
        fileInfo.textContent = 'Error: Parsing failed. Invalid JSON.';
        return;
      }

      try {
        this.renderHouse(data);
        fileInfo.textContent = `Applied: ${file.name}`;

        const statsPanel = document.getElementById('stats-panel') as HTMLElement;
        const wallCount = document.getElementById('wall-count') as HTMLElement;
        statsPanel.style.display = 'block';
        wallCount.textContent = data.walls.length.toString();
      } catch (e) {
        fileInfo.textContent = 'Error: Rendering failed.';
        console.error('Rendering error:', e);
      }
    } catch (error) {
      console.error('File reading error:', error);
      fileInfo.textContent = 'Error: Could not read file.';
    }
  }

  private renderHouse(data: JsonData) {
    // Clear existing walls
    while (this.wallsGroup.children.length > 0) {
      const child = this.wallsGroup.children[0] as THREE.Mesh;
      child.geometry.dispose();
      (child.material as THREE.Material).dispose();
      this.wallsGroup.remove(child);
    }

    if (!data.walls || data.walls.length === 0) return;

    // Calculate bounding box of all walls to Center & Scale
    let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
    data.walls.forEach(wall => {
      minX = Math.min(minX, wall.bbox.x1, wall.bbox.x2);
      maxX = Math.max(maxX, wall.bbox.x1, wall.bbox.x2);
      minY = Math.min(minY, wall.bbox.y1, wall.bbox.y2);
      maxY = Math.max(maxY, wall.bbox.y1, wall.bbox.y2);
    });

    const centerX = (minX + maxX) / 2;
    const centerY = (minY + maxY) / 2;
    const size = Math.max(maxX - minX, maxY - minY);
    const scale = 50 / size; // Scale to fit in ~50 units

    data.walls.forEach(wall => {
      const width = Math.abs(wall.bbox.x2 - wall.bbox.x1) * scale;
      const depth = Math.abs(wall.bbox.y2 - wall.bbox.y1) * scale;

      // Three.js Coordinate System:
      // X = JSON X
      // Y = Height (Up)
      // Z = JSON Y

      const geometry = new THREE.BoxGeometry(
        width || 0.1, // Minimum width if 0
        this.wallHeight,
        depth || 0.1  // Minimum depth if 0
      );

      const material = this.wallMaterials[wall.class] || this.wallMaterials.default;
      const mesh = new THREE.Mesh(geometry, material);

      // Position (compensate for Three.js centering box)
      const posX = (wall.center.x - centerX) * scale;
      const posZ = (wall.center.y - centerY) * scale;

      mesh.position.set(posX, this.wallHeight / 2, posZ);
      this.wallsGroup.add(mesh);
    });

    // Reset camera to view the whole house
    this.camera.position.set(40, 40, 40);
    this.controls.target.set(0, 0, 0);
    this.controls.update();
  }

  private animate() {
    requestAnimationFrame(() => this.animate());
    this.controls.update();
    this.renderer.render(this.scene, this.camera);
  }
}

new HouseViewer();
