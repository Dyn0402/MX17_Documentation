#!/usr/bin/env python3
"""
Overhead (top-down) schematic sketches of the n_TOF EAR2 X17 beam-test setups.
Beam (proton bunch -> Pb spallation target 20 m below) arrives along +Y, i.e.
INTO the page.  The plane drawn is the horizontal XZ plane.

  Figure A - February 2026:  single MX17 Micromegas + one SiPM wall behind it.
  Figure B - May 2026:        two MX17 Micromegas at 90 deg.
                              Det-3 arm:  MM -> SiPM wall -> GDD scintillator
                              Det-4 arm:  MM -> GDD scintillator

Dimensions (transverse sizes, cm) taken from the beam-test documentation:
  MX17 Micromegas active area : 40 x 40 cm
  SiPM wall                   : 50 x 50 cm (20 x 2.5 cm bars)
  GDD trigger scintillator    : 30 x 30 cm
  Front face of MM            : ~20 cm from beam axis
"""
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import Rectangle, FancyArrow

# ---- colours (shared with plot_geometry.py palette) -------------------------
C_MM    = '#4a90d9'   # Micromegas
C_SIPM  = '#e07820'   # SiPM wall
C_SCINT = '#f0c040'   # plastic / GDD scintillator
C_BEAM  = '#c0392b'

# ---- transverse half-widths [cm] --------------------------------------------
HW = dict(mm=40/2, sipm=50/2, scint=30/2)
# ---- stack depths (drawn, exaggerated for visibility) [cm] ------------------
T  = dict(mm=3.0, sipm=4.0, scint=3.0)
GAP = 4.0                 # drawn gap between elements
MM_FRONT = 20.0           # MM front face distance from beam axis [cm]


def slab(ax, arm_dir, w_front, half_width, depth, color, label, z=2):
    """Draw a slab in the overhead plane.
    arm_dir : unit vector (x,z) pointing from beam axis outward (stack axis).
    The transverse (width) axis is perpendicular to arm_dir in-plane."""
    d = np.array(arm_dir, float); d /= np.linalg.norm(d)
    perp = np.array([-d[1], d[0]])
    w_back = w_front + depth
    centre = d * (w_front + w_back) / 2.0
    # rectangle corners
    corners = []
    for sgn_w, sgn_p in [(-1,-1),(-1,1),(1,1),(1,-1)]:
        pt = (d * (w_front if sgn_w<0 else w_back)
              + perp * sgn_p * half_width)
        corners.append(pt)
    poly = plt.Polygon(corners, closed=True, facecolor=color,
                       edgecolor='k', linewidth=1.0, alpha=0.85, zorder=z)
    ax.add_patch(poly)
    return centre, perp, d


def beam_marker(ax):
    ax.add_patch(plt.Circle((0,0), 1.6, facecolor='white',
                            edgecolor=C_BEAM, linewidth=2.2, zorder=6))
    ax.add_patch(plt.Circle((0,0), 0.45, facecolor=C_BEAM, zorder=7))
    ax.annotate('beam $\\odot$ (+Y, into page)\n$\\gamma$-flash source',
                xy=(0,0), xytext=(2.6, 2.6), fontsize=9,
                color=C_BEAM, fontweight='bold', zorder=7)


def dim_arrow(ax, p0, p1, label, off=0):
    p0=np.array(p0,float); p1=np.array(p1,float)
    ax.annotate('', xy=p1, xytext=p0,
                arrowprops=dict(arrowstyle='<->', color='0.35', lw=1.2))
    mid=(p0+p1)/2
    ax.text(mid[0], mid[1]-1.6-off, label, ha='center', va='top',
            fontsize=8, color='0.3')


# =============================================================================
# Figure A - February
# =============================================================================
def fig_february():
    fig, ax = plt.subplots(figsize=(9.5, 7.2))
    ax.set_aspect('equal')
    d = (1, 0)                      # arm points +X
    perp = np.array([0,1.0])

    w = MM_FRONT
    slab(ax, d, w, HW['mm'],   T['mm'],   C_MM,   'MM');   w += T['mm']+GAP
    slab(ax, d, w, HW['sipm'], T['sipm'], C_SIPM, 'SiPM'); w_sipm=w; w += T['sipm']

    # element labels
    ax.text(MM_FRONT+T['mm']/2, HW['mm']+2.5, 'MX17\nMicromegas\n(40×40 cm)',
            ha='center', va='bottom', fontsize=9, fontweight='bold')
    ax.text(w_sipm+T['sipm']/2, -HW['sipm']-2.5, 'SiPM wall\n(50×50 cm)',
            ha='center', va='top', fontsize=9, fontweight='bold', color='#a85a10')

    beam_marker(ax)
    dim_arrow(ax, (0,-HW['sipm']-9), (MM_FRONT,-HW['sipm']-9), '≈20 cm')

    lim=42
    ax.set_xlim(-8, lim); ax.set_ylim(-lim+4, lim-4)
    ax.set_xlabel('X  [cm]'); ax.set_ylabel('Z  [cm]')
    ax.set_title('February 2026 — Overhead view\nSingle Micromegas + SiPM wall',
                 fontsize=12, fontweight='bold')
    ax.axhline(0,color='0.85',lw=0.6,zorder=0); ax.axvline(0,color='0.85',lw=0.6,zorder=0)
    ax.grid(True, lw=0.3, alpha=0.4)
    leg=[mpatches.Patch(color=C_MM,label='MX17 Micromegas (40×40 cm)'),
         mpatches.Patch(color=C_SIPM,label='SiPM wall (50×50 cm, 20 bars)')]
    ax.legend(handles=leg, loc='upper right', fontsize=8.5, framealpha=0.9)
    fig.tight_layout()
    fig.savefig('assets/setup_february.png', dpi=200, bbox_inches='tight')
    print('saved assets/setup_february.png')


# =============================================================================
# Figure B - May
# =============================================================================
def fig_may():
    fig, ax = plt.subplots(figsize=(9.8, 8.6))
    ax.set_aspect('equal')

    # ---- Det-3 arm: +X  (MM -> SiPM wall -> GDD scint) ----
    d3=(1,0)
    w=MM_FRONT
    slab(ax,d3,w,HW['mm'],T['mm'],C_MM,'MM'); w+=T['mm']+GAP
    slab(ax,d3,w,HW['sipm'],T['sipm'],C_SIPM,'SiPM'); w_s3=w; w+=T['sipm']+GAP
    slab(ax,d3,w,HW['scint'],T['scint'],C_SCINT,'Scint'); w_sc3=w
    ax.text(MM_FRONT+T['mm']/2, HW['mm']+2.2,'MX17-3\n(Det-3)\n40×40 cm',
            ha='center',va='bottom',fontsize=8.5,fontweight='bold')
    ax.text(w_s3+T['sipm']/2,-HW['sipm']-2.2,'SiPM wall\n50×50 cm',
            ha='center',va='top',fontsize=8,fontweight='bold',color='#a85a10')
    ax.text(w_sc3+T['scint']/2, HW['scint']+2.2,'GDD scint\n30×30 cm',
            ha='center',va='bottom',fontsize=8,color='#9a7d10')

    # ---- Det-4 arm: -Z  (MM -> GDD scint) ----
    d4=(0,-1)
    w=MM_FRONT
    slab(ax,d4,w,HW['mm'],T['mm'],C_MM,'MM'); w+=T['mm']+GAP
    slab(ax,d4,w,HW['scint'],T['scint'],C_SCINT,'Scint'); w_sc4=w
    ax.text(-HW['mm']-2.2, -(MM_FRONT+T['mm']/2),'MX17-4\n(Det-4)\n40×40 cm',
            ha='right',va='center',fontsize=8.5,fontweight='bold')
    ax.text(HW['scint']+2.2, -(w_sc4+T['scint']/2),'GDD scint\n30×30 cm',
            ha='left',va='center',fontsize=8,color='#9a7d10')

    # 90 degree marker
    ax.annotate('', xy=(13,0), xytext=(0,0),
                arrowprops=dict(arrowstyle='-',color='0.5',lw=0.8,ls='--'))
    ax.annotate('', xy=(0,-13), xytext=(0,0),
                arrowprops=dict(arrowstyle='-',color='0.5',lw=0.8,ls='--'))
    ax.add_patch(mpatches.Arc((0,0),20,20,theta1=-90,theta2=0,color='0.5',lw=1.0))
    ax.text(9.7,-9.7,'90°',fontsize=10,color='0.4',fontweight='bold')

    beam_marker(ax)
    dim_arrow(ax,(0,HW['sipm']+9),(MM_FRONT,HW['sipm']+9),'≈20 cm')

    lim=46
    ax.set_xlim(-lim+6, lim); ax.set_ylim(-lim, lim-6)
    ax.set_xlabel('X  [cm]'); ax.set_ylabel('Z  [cm]')
    ax.set_title('May 2026 — Overhead view\nTwo Micromegas at 90°  ·  SiPM wall behind Det-3, '
                 'GDD scintillator behind each',fontsize=11.5, fontweight='bold')
    ax.axhline(0,color='0.85',lw=0.6,zorder=0); ax.axvline(0,color='0.85',lw=0.6,zorder=0)
    ax.grid(True, lw=0.3, alpha=0.4)
    leg=[mpatches.Patch(color=C_MM,label='MX17 Micromegas (40×40 cm)'),
         mpatches.Patch(color=C_SIPM,label='SiPM wall (50×50 cm, 20 bars)'),
         mpatches.Patch(color=C_SCINT,label='GDD trigger scintillator (30×30 cm)')]
    ax.legend(handles=leg, loc='lower left', fontsize=8.5, framealpha=0.9)
    fig.tight_layout()
    fig.savefig('assets/setup_may.png', dpi=200, bbox_inches='tight')
    print('saved assets/setup_may.png')


if __name__=='__main__':
    fig_february()
    fig_may()
