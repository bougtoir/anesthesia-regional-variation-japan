#!/usr/bin/env python3
"""
Generate 3D extruded polygon maps v2 - ENGLISH labels.
Rebuilt from scratch:
- All 335 SMAs rendered (missing-data areas shown flat + neutral colour)
- area_code zero-padded to 4 digits so Hokkaido/Tohoku/Kanto merge correctly
- No lat/lon axis labels
- North at top, near-overhead camera
"""

import geopandas as gpd
import pandas as pd
import numpy as np
from scipy.spatial import Delaunay
import plotly.graph_objects as go
import gc
import os
import time

print("Loading data...")
gdf = gpd.read_file('/home/ubuntu/gis_data/merged_enriched.gpkg')
gdf_nt = gpd.read_file('/home/ubuntu/gis_data/northern_territories.gpkg')
corr = pd.read_csv('/home/ubuntu/corrected_metrics_final.csv')
anes = pd.read_csv('/home/ubuntu/anesthesiologist_by_sma.csv')

# ── KEY FIX: zero-pad area_code to 4 digits before merge ──
anes.rename(columns={'code': 'area_code', 'anesthesiologist_count': 'n_anes'}, inplace=True)
gdf['area_code'] = gdf['area_code'].astype(str).str.zfill(4)
corr['area_code'] = corr['area_code'].astype(str).str.zfill(4)
anes['area_code'] = anes['area_code'].astype(str).str.zfill(4)

gdf = gdf.merge(corr[['area_code', 'L008_per_surgery', 'L003_per_surgery', 'L003_L008_ratio']],
                 on='area_code', how='left', suffixes=('', '_corr'))
gdf = gdf.merge(anes[['area_code', 'n_anes']], on='area_code', how='left')
if 'L003_L008_ratio_corr' in gdf.columns:
    gdf['L003_L008_ratio'] = gdf['L003_L008_ratio_corr'].fillna(gdf['L003_L008_ratio'])

# Simplify geometries
print("Simplifying geometries...")
gdf['geometry'] = gdf.geometry.simplify(tolerance=0.03, preserve_topology=True)

# Simplify Northern Territories (many tiny islands)
print("Simplifying Northern Territories...")
gdf_nt['geometry'] = gdf_nt.geometry.simplify(tolerance=0.01, preserve_topology=True)

# Verify data coverage
n_ratio = gdf['L003_L008_ratio'].notna().sum()
n_anes = gdf['n_anes'].notna().sum()
print(f"Total: {len(gdf)} areas | L003_L008_ratio: {n_ratio} | n_anes: {n_anes}")

os.makedirs('/home/ubuntu/3d_extruded', exist_ok=True)


def polygon_to_triangles(coords_2d):
    """Triangulate a 2D polygon using Delaunay, filtering exterior triangles."""
    from shapely.geometry import Point, Polygon as ShapelyPolygon
    poly = ShapelyPolygon(coords_2d)
    if not poly.is_valid or poly.area < 1e-8:
        return np.array([], dtype=int).reshape(0, 3)
    pts = np.array(coords_2d[:-1])
    if len(pts) < 3:
        return np.array([], dtype=int).reshape(0, 3)
    try:
        tri = Delaunay(pts)
    except Exception:
        return np.array([], dtype=int).reshape(0, 3)
    valid = []
    for simplex in tri.simplices:
        centroid = pts[simplex].mean(axis=0)
        if poly.contains(Point(centroid)):
            valid.append(simplex)
    if not valid:
        return np.array([], dtype=int).reshape(0, 3)
    return np.array(valid)


def build_mesh_data(gdf_all, color_col, height_col, height_scale):
    """Build mesh vertex/triangle data for ALL areas.
    Areas with missing colour/height rendered flat (h=0) at neutral colour.
    """
    all_x, all_y, all_z = [], [], []
    all_i, all_j, all_k = [], [], []
    all_intensity = []
    offset = 0
    count = 0

    # Neutral fill for missing data
    valid_colors = gdf_all[color_col].dropna()
    color_mid = valid_colors.median() if len(valid_colors) > 0 else 0

    for idx, row in gdf_all.iterrows():
        geom = row.geometry
        if geom is None:
            continue
        color_val = row[color_col] if pd.notna(row[color_col]) else color_mid
        h_raw = row[height_col] if pd.notna(row[height_col]) else 0
        h = max(0, h_raw) * height_scale

        polys = []
        if geom.geom_type == 'Polygon':
            polys = [geom]
        elif geom.geom_type == 'MultiPolygon':
            polys = list(geom.geoms)

        for poly in polys:
            coords = np.array(poly.exterior.coords)
            if len(coords) < 4:
                continue

            pts_2d = coords[:-1]
            n = len(pts_2d)
            if n < 3:
                continue

            triangles = polygon_to_triangles(coords)
            if len(triangles) == 0:
                continue

            # Top face
            for pt in pts_2d:
                all_x.append(pt[0])
                all_y.append(pt[1])
                all_z.append(h)
                all_intensity.append(color_val)

            for tri in triangles:
                all_i.append(tri[0] + offset)
                all_j.append(tri[1] + offset)
                all_k.append(tri[2] + offset)

            # Bottom face
            bottom_offset = offset + n
            for pt in pts_2d:
                all_x.append(pt[0])
                all_y.append(pt[1])
                all_z.append(0)
                all_intensity.append(color_val)

            for tri in triangles:
                all_i.append(tri[2] + bottom_offset)
                all_j.append(tri[1] + bottom_offset)
                all_k.append(tri[0] + bottom_offset)

            # Side faces
            step = max(1, n // 25)
            for si in range(0, n, step):
                sj = (si + step) % n
                t_si = offset + si
                t_sj = offset + sj
                b_si = bottom_offset + si
                b_sj = bottom_offset + sj
                all_i.append(t_si)
                all_j.append(b_si)
                all_k.append(b_sj)
                all_i.append(t_si)
                all_j.append(b_sj)
                all_k.append(t_sj)

            offset += 2 * n
            count += 1

    return all_x, all_y, all_z, all_i, all_j, all_k, all_intensity, count


def build_nt_mesh(gdf_nt):
    """Build flat mesh for Northern Territories (white, height=0)."""
    all_x, all_y, all_z = [], [], []
    all_i, all_j, all_k = [], [], []
    offset = 0
    count = 0

    for idx, row in gdf_nt.iterrows():
        geom = row.geometry
        if geom is None:
            continue
        polys = []
        if geom.geom_type == 'Polygon':
            polys = [geom]
        elif geom.geom_type == 'MultiPolygon':
            polys = list(geom.geoms)

        for poly in polys:
            coords = np.array(poly.exterior.coords)
            if len(coords) < 4:
                continue
            pts_2d = coords[:-1]
            n = len(pts_2d)
            if n < 3:
                continue
            triangles = polygon_to_triangles(coords)
            if len(triangles) == 0:
                continue
            for pt in pts_2d:
                all_x.append(pt[0])
                all_y.append(pt[1])
                all_z.append(0)
            for tri in triangles:
                all_i.append(tri[0] + offset)
                all_j.append(tri[1] + offset)
                all_k.append(tri[2] + offset)
            offset += n
            count += 1

    return all_x, all_y, all_z, all_i, all_j, all_k, count


def make_3d_map_v2(gdf_all, gdf_nt, color_col, height_col, title, output_html,
                   color_scale='PiYG', color_vmin=0, color_vmax=3,
                   height_scale=0.01,
                   color_label=''):
    """Create 3D extruded map with all 335 SMAs + Northern Territories, north at top."""
    t0 = time.time()

    # Build mesh for all SMA areas
    print("  Building SMA mesh...")
    all_x, all_y, all_z, all_i, all_j, all_k, all_intensity, total_c = build_mesh_data(
        gdf_all, color_col, height_col, height_scale)

    # Build Northern Territories mesh
    print("  Building Northern Territories mesh...")
    nt_x, nt_y, nt_z, nt_i, nt_j, nt_k, nt_count = build_nt_mesh(gdf_nt)

    elapsed = time.time() - t0
    print(f"  Total mesh: {total_c} SMA polygons + {nt_count} NT polygons ({elapsed:.1f}s)")

    mesh = go.Mesh3d(
        x=all_x, y=all_y, z=all_z,
        i=all_i, j=all_j, k=all_k,
        intensity=all_intensity,
        colorscale=color_scale,
        cmin=color_vmin, cmax=color_vmax,
        colorbar=dict(
            title=dict(text=color_label, font=dict(size=14)),
            len=0.6, thickness=20,
            x=1.02,
        ),
        flatshading=True,
        lighting=dict(ambient=0.6, diffuse=0.5, specular=0.1, roughness=0.5),
        lightposition=dict(x=0, y=0, z=100),
    )

    # Northern Territories: flat white mesh (no colorbar)
    nt_mesh = go.Mesh3d(
        x=nt_x, y=nt_y, z=nt_z,
        i=nt_i, j=nt_j, k=nt_k,
        color='#e8e8e8',
        flatshading=True,
        lighting=dict(ambient=0.8, diffuse=0.3, specular=0.0),
        showscale=False,
        hoverinfo='skip',
    )

    fig = go.Figure(data=[mesh, nt_mesh])

    # Camera: looking from south (negative y) toward north, with elevation
    # In plotly: eye.y negative = looking from south, eye.z positive = looking down slightly
    fig.update_layout(
        title=dict(text=title, font=dict(size=16), x=0.5),
        scene=dict(
            # Hide all axis labels and tick labels
            xaxis=dict(
                showticklabels=False, title='', showgrid=False,
                showline=False, zeroline=False, showbackground=False,
            ),
            yaxis=dict(
                showticklabels=False, title='', showgrid=False,
                showline=False, zeroline=False, showbackground=False,
            ),
            zaxis=dict(
                showticklabels=False, title='', showgrid=False,
                showline=False, zeroline=False, showbackground=False,
            ),
            aspectmode='manual',
            aspectratio=dict(x=1.8, y=1.5, z=0.08),
            camera=dict(
                # Near-overhead view (~80° elevation, ~10° from top-down)
                # Raised high enough to see all of Japan (Okinawa to Hokkaido)
                eye=dict(x=0.0, y=-0.4, z=2.5),
                up=dict(x=0, y=1, z=0),
                center=dict(x=0, y=0, z=0),
            ),
        ),
        width=1800, height=1400,
        margin=dict(l=0, r=80, t=60, b=0),
        paper_bgcolor='white',
        plot_bgcolor='white',
    )

    fig.write_html(output_html, include_plotlyjs='cdn')
    print(f"  Saved: {output_html}")
    gc.collect()


configs = [
    {
        'name': 'ratio_by_anes',
        'color_col': 'L003_L008_ratio',
        'height_col': 'n_anes',
        'title': 'GA+Regional / GA Ratio (colour) × Anaesthesiologist Count (height)',
        'color_scale': 'PiYG', 'color_vmin': 0, 'color_vmax': 3,
        'height_scale': 0.01,
        'color_label': 'L003/L008 Ratio',
        'height_label': 'Anes. Count',
    },
    {
        'name': 'L008_by_anes',
        'color_col': 'L008_scr',
        'height_col': 'n_anes',
        'title': 'GA SCR (colour) × Anaesthesiologist Count (height)',
        'color_scale': 'RdYlBu_r', 'color_vmin': 0, 'color_vmax': 250,
        'height_scale': 0.01,
        'color_label': 'L008 SCR',
        'height_label': 'Anes. Count',
    },
    {
        'name': 'ratio_by_surgery',
        'color_col': 'L003_L008_ratio',
        'height_col': 'L008_per_surgery',
        'title': 'GA+Regional / GA Ratio (colour) × GA/Surgery Rate (height)',
        'color_scale': 'PiYG', 'color_vmin': 0, 'color_vmax': 3,
        'height_scale': 3.0,
        'color_label': 'L003/L008 Ratio',
        'height_label': 'L008/Surgery',
    },
]

for i, cfg in enumerate(configs, 1):
    print(f"\n[{i}/{len(configs)}] Generating {cfg['name']}...")
    make_3d_map_v2(
        gdf, gdf_nt,
        cfg['color_col'], cfg['height_col'],
        cfg['title'],
        f"/home/ubuntu/3d_extruded/3D_{cfg['name']}_v2.html",
        color_scale=cfg['color_scale'],
        color_vmin=cfg['color_vmin'], color_vmax=cfg['color_vmax'],
        height_scale=cfg['height_scale'],
        color_label=cfg['color_label'],
    )

print("\nAll EN v2 maps generated!")
