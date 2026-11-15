// Vector path building shared by every graphics source (DrawingML geometry, VML, charts, EMF/WMF).
//
// Paths are kept as absolute move/line/cubic/close commands only: quadratic curves are raised to
// cubics and elliptical arcs are split into cubic Béziers when added. That keeps every affine
// transform exact (just map the points), which is how shape rotation/flips, group child
// coordinate spaces and metafile window/viewport mappings are applied. Output is SVG path data
// with numbers rounded to 1/1000 pt, so the same input always gives the same string.

/** Affine matrix [a, b, c, d, e, f]: x' = a·x + c·y + e, y' = b·x + d·y + f (SVG/PDF order). */
export type Matrix = readonly [number, number, number, number, number, number];

export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/** `m` then `n` (apply m first). */
export function multiply(m: Matrix, n: Matrix): Matrix {
  return [
    m[0] * n[0] + m[1] * n[2],
    m[0] * n[1] + m[1] * n[3],
    m[2] * n[0] + m[3] * n[2],
    m[2] * n[1] + m[3] * n[3],
    m[4] * n[0] + m[5] * n[2] + n[4],
    m[4] * n[1] + m[5] * n[3] + n[5],
  ];
}

export const translate = (x: number, y: number): Matrix => [1, 0, 0, 1, x, y];
export const scale = (sx: number, sy: number): Matrix => [sx, 0, 0, sy, 0, 0];
/** Rotation by `deg` degrees, clockwise on screen (y grows downward). */
export function rotate(deg: number): Matrix {
  const r = (deg * Math.PI) / 180;
  const c = Math.cos(r);
  const s = Math.sin(r);
  return [c, s, -s, c, 0, 0];
}

/** Maps a point. */
export function apply(m: Matrix, x: number, y: number): [number, number] {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

/** Geometric mean scale of a matrix (stroke widths under a transform). */
export function matrixScale(m: Matrix): number {
  return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2]));
}

/**
 * Transform placing a box of size (w, h) whose local origin is its top-left at (x, y), rotated
 * `rot` degrees about its centre and flipped about its centre axes first.
 */
export function boxTransform(
  x: number,
  y: number,
  w: number,
  h: number,
  rot: number,
  flipH: boolean,
  flipV: boolean,
): Matrix {
  const cx = w / 2;
  const cy = h / 2;
  let m: Matrix = translate(-cx, -cy);
  if (flipH || flipV) m = multiply(m, scale(flipH ? -1 : 1, flipV ? -1 : 1));
  if (rot !== 0) m = multiply(m, rotate(rot));
  return multiply(m, translate(x + cx, y + cy));
}

/** Rounds to 1/1000 and prints without exponent or trailing zeros ("-0" → "0"). */
export function fmt(n: number): string {
  if (!Number.isFinite(n)) return '0';
  const r = Math.round(n * 1000) / 1000;
  return r === 0 ? '0' : String(r);
}

/** Path commands: M/L carry one point, C three, Z none. */
type Cmd = { op: 'M' | 'L'; p: number[] } | { op: 'C'; p: number[] } | { op: 'Z'; p: number[] };

/** A path under construction (absolute coordinates). */
export class Path {
  readonly cmds: Cmd[] = [];
  /** Current point. */
  x = 0;
  y = 0;
  #startX = 0;
  #startY = 0;
  #open = false;

  get empty(): boolean {
    return this.cmds.length === 0;
  }

  moveTo(x: number, y: number): this {
    this.cmds.push({ op: 'M', p: [x, y] });
    this.x = this.#startX = x;
    this.y = this.#startY = y;
    this.#open = true;
    return this;
  }

  lineTo(x: number, y: number): this {
    if (!this.#open) this.moveTo(this.x, this.y);
    this.cmds.push({ op: 'L', p: [x, y] });
    this.x = x;
    this.y = y;
    return this;
  }

  cubicTo(x1: number, y1: number, x2: number, y2: number, x: number, y: number): this {
    if (!this.#open) this.moveTo(this.x, this.y);
    this.cmds.push({ op: 'C', p: [x1, y1, x2, y2, x, y] });
    this.x = x;
    this.y = y;
    return this;
  }

  quadTo(qx: number, qy: number, x: number, y: number): this {
    const x0 = this.x;
    const y0 = this.y;
    return this.cubicTo(
      x0 + ((qx - x0) * 2) / 3,
      y0 + ((qy - y0) * 2) / 3,
      x + ((qx - x) * 2) / 3,
      y + ((qy - y) * 2) / 3,
      x,
      y,
    );
  }

  /**
   * Elliptical arc of the axis-aligned ellipse with centre (cx, cy) and radii (rx, ry), from
   * parametric angle `t0` sweeping `dt` radians (positive = clockwise on screen). Draws a line to
   * the start point when it is not the current point (or moves there when nothing is open).
   */
  ellipseArc(cx: number, cy: number, rx: number, ry: number, t0: number, dt: number): this {
    const sx = cx + rx * Math.cos(t0);
    const sy = cy + ry * Math.sin(t0);
    if (!this.#open) this.moveTo(sx, sy);
    else if (Math.abs(sx - this.x) > 1e-9 || Math.abs(sy - this.y) > 1e-9) this.lineTo(sx, sy);
    if (dt === 0) return this;
    const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9));
    const step = dt / n;
    const k = (4 / 3) * Math.tan(step / 4);
    let a = t0;
    for (let i = 0; i < n; i++) {
      const b = a + step;
      const ca = Math.cos(a);
      const sa = Math.sin(a);
      const cb = Math.cos(b);
      const sb = Math.sin(b);
      this.cubicTo(
        cx + rx * (ca - k * sa),
        cy + ry * (sa + k * ca),
        cx + rx * (cb + k * sb),
        cy + ry * (sb - k * cb),
        cx + rx * cb,
        cy + ry * sb,
      );
      a = b;
    }
    return this;
  }

  /**
   * DrawingML `arcTo`: the current point lies on an ellipse with radii (wR, hR) at visual angle
   * `stAng`; sweep `swAng` (radians, clockwise). Visual angles are converted to parametric ones.
   */
  arcTo(wR: number, hR: number, stAng: number, swAng: number): this {
    if (wR === 0 || hR === 0) {
      // Degenerate ellipse: the arc collapses onto a line through the end point.
      const end = stAng + swAng;
      const x = this.x - wR * Math.cos(stAng) + wR * Math.cos(end);
      const y = this.y - hR * Math.sin(stAng) + hR * Math.sin(end);
      return this.lineTo(x, y);
    }
    const t1 = visualToParam(stAng, wR, hR);
    let dt: number;
    if (Math.abs(swAng) >= 2 * Math.PI - 1e-9) dt = Math.sign(swAng) * 2 * Math.PI;
    else {
      dt = visualToParam(stAng + swAng, wR, hR) - t1;
      while (swAng > 0 && dt < 0) dt += 2 * Math.PI;
      while (swAng < 0 && dt > 0) dt -= 2 * Math.PI;
      if (swAng === 0) dt = 0;
    }
    const cx = this.x - wR * Math.cos(t1);
    const cy = this.y - hR * Math.sin(t1);
    return this.ellipseArc(cx, cy, wR, hR, t1, dt);
  }

  close(): this {
    if (this.#open) {
      this.cmds.push({ op: 'Z', p: [] });
      this.x = this.#startX;
      this.y = this.#startY;
      this.#open = false;
    }
    return this;
  }

  /** Full ellipse in the box (x, y, w, h). */
  ellipse(x: number, y: number, w: number, h: number): this {
    const rx = w / 2;
    const ry = h / 2;
    this.moveTo(x + w, y + ry);
    this.ellipseArc(x + rx, y + ry, rx, ry, 0, 2 * Math.PI);
    return this.close();
  }

  rect(x: number, y: number, w: number, h: number): this {
    return this.moveTo(x, y)
      .lineTo(x + w, y)
      .lineTo(x + w, y + h)
      .lineTo(x, y + h)
      .close();
  }

  /** Polygon or polyline through `pts` ([x0, y0, x1, y1, ...]). */
  poly(pts: readonly number[], closed: boolean): this {
    if (pts.length < 2) return this;
    this.moveTo(pts[0] ?? 0, pts[1] ?? 0);
    for (let i = 2; i + 1 < pts.length; i += 2) this.lineTo(pts[i] ?? 0, pts[i + 1] ?? 0);
    return closed ? this.close() : this;
  }

  /** Appends another path's commands. */
  append(other: Path): this {
    for (const c of other.cmds) this.cmds.push({ op: c.op, p: c.p.slice() } as Cmd);
    this.x = other.x;
    this.y = other.y;
    this.#open = false;
    return this;
  }

  /** A transformed copy. */
  transform(m: Matrix): Path {
    const out = new Path();
    for (const c of this.cmds) {
      const p: number[] = [];
      for (let i = 0; i + 1 < c.p.length; i += 2) {
        const [x, y] = apply(m, c.p[i] ?? 0, c.p[i + 1] ?? 0);
        p.push(x, y);
      }
      out.cmds.push({ op: c.op, p } as Cmd);
    }
    return out;
  }

  /** True when every subpath ends with a close. */
  get closed(): boolean {
    let open = false;
    for (const c of this.cmds) {
      if (c.op === 'M') {
        if (open) return false;
        open = true;
      } else if (c.op === 'Z') open = false;
    }
    return !open;
  }

  /**
   * The area to fill: open subpaths are closed implicitly (as DrawingML, VML and PDF fill them)
   * and subpaths without area (a lone point or a single line segment) are dropped.
   */
  fillable(): Path {
    if (this.closed) return this;
    const out = new Path();
    let sub: Cmd[] = [];
    const flush = (): void => {
      const area = sub.length > 2 || sub.some((c) => c.op === 'C');
      if (area) {
        for (const c of sub) out.cmds.push(c);
        if (sub[sub.length - 1]?.op !== 'Z') out.cmds.push({ op: 'Z', p: [] });
      }
      sub = [];
    };
    for (const c of this.cmds) {
      if (c.op === 'M') flush();
      sub.push(c);
      if (c.op === 'Z') flush();
    }
    flush();
    return out;
  }

  /** Bounding box of the control points [minX, minY, maxX, maxY] (undefined when empty). */
  bounds(): [number, number, number, number] | undefined {
    let x0 = Infinity;
    let y0 = Infinity;
    let x1 = -Infinity;
    let y1 = -Infinity;
    for (const c of this.cmds) {
      for (let i = 0; i + 1 < c.p.length; i += 2) {
        const x = c.p[i] ?? 0;
        const y = c.p[i + 1] ?? 0;
        if (x < x0) x0 = x;
        if (x > x1) x1 = x;
        if (y < y0) y0 = y;
        if (y > y1) y1 = y;
      }
    }
    return x0 === Infinity ? undefined : [x0, y0, x1, y1];
  }

  /** Subpaths flattened to polylines (curves subdivided), each with its closed flag. */
  flatten(): { pts: number[]; closed: boolean }[] {
    const out: { pts: number[]; closed: boolean }[] = [];
    let cur: { pts: number[]; closed: boolean } | undefined;
    let x = 0;
    let y = 0;
    for (const c of this.cmds) {
      if (c.op === 'M') {
        cur = { pts: [c.p[0] ?? 0, c.p[1] ?? 0], closed: false };
        out.push(cur);
        x = c.p[0] ?? 0;
        y = c.p[1] ?? 0;
      } else if (c.op === 'L') {
        cur?.pts.push(c.p[0] ?? 0, c.p[1] ?? 0);
        x = c.p[0] ?? 0;
        y = c.p[1] ?? 0;
      } else if (c.op === 'C') {
        const [x1 = 0, y1 = 0, x2 = 0, y2 = 0, x3 = 0, y3 = 0] = c.p;
        const n = 12;
        for (let i = 1; i <= n; i++) {
          const t = i / n;
          const u = 1 - t;
          cur?.pts.push(
            u * u * u * x + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t * t * t * x3,
            u * u * u * y + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t * t * t * y3,
          );
        }
        x = x3;
        y = y3;
      } else if (cur !== undefined) {
        cur.closed = true;
        x = cur.pts[0] ?? 0;
        y = cur.pts[1] ?? 0;
      }
    }
    return out;
  }

  /** SVG path data. */
  toString(): string {
    let s = '';
    for (const c of this.cmds) {
      if (s !== '') s += ' ';
      s += c.op;
      for (const v of c.p) s += ` ${fmt(v)}`;
    }
    return s;
  }
}

/** Visual (true) angle on an ellipse → parametric angle. */
function visualToParam(ang: number, rx: number, ry: number): number {
  return Math.atan2(rx * Math.sin(ang), ry * Math.cos(ang));
}

/**
 * A dashed version of `path` as open stroke segments (PathItem has no dash array, so dashes are
 * materialised). `pattern` alternates dash/gap lengths in points.
 */
export function dashPath(path: Path, pattern: readonly number[]): Path {
  const out = new Path();
  const total = pattern.reduce((a, b) => a + b, 0);
  if (pattern.length < 2 || !(total > 0)) return path;
  for (const sub of path.flatten()) {
    const pts = sub.pts.slice();
    if (sub.closed && pts.length >= 2) pts.push(pts[0] ?? 0, pts[1] ?? 0);
    let k = 0; // index into pattern
    let left = pattern[0] ?? 0; // remaining length of the current dash/gap
    let on = true;
    let drawing = false;
    for (let i = 0; i + 3 < pts.length; i += 2) {
      let x0 = pts[i] ?? 0;
      let y0 = pts[i + 1] ?? 0;
      const x1 = pts[i + 2] ?? 0;
      const y1 = pts[i + 3] ?? 0;
      let seg = Math.hypot(x1 - x0, y1 - y0);
      while (seg > 0) {
        const step = Math.min(seg, left);
        const t = step / seg;
        const nx = x0 + (x1 - x0) * t;
        const ny = y0 + (y1 - y0) * t;
        if (on) {
          if (!drawing) out.moveTo(x0, y0);
          out.lineTo(nx, ny);
          drawing = true;
        }
        x0 = nx;
        y0 = ny;
        seg -= step;
        left -= step;
        if (left <= 1e-9) {
          k = (k + 1) % pattern.length;
          left = pattern[k] ?? 0;
          on = k % 2 === 0;
          drawing = false;
        }
      }
    }
  }
  return out;
}
