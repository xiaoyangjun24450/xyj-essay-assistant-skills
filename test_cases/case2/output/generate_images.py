#!/usr/bin/env python3
"""
Generate images for ESP32 FOC Controller Design document
"""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import os

# Create images directory
os.makedirs('images', exist_ok=True)

# Set style
plt.style.use('seaborn-v0_8-darkgrid')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['font.size'] = 10

# ============================================================
# 1. FOC Control System Architecture Block Diagram
# ============================================================
fig, ax = plt.subplots(figsize=(12, 6))
ax.set_xlim(0, 10)
ax.set_ylim(0, 6)
ax.axis('off')

# Define box positions and labels
boxes = [
    (0.5, 3, 1.5, 'Current\nSensor'),
    (2.5, 3, 1.5, 'Clarke\nTransform'),
    (4.5, 3, 1.5, 'Park\nTransform'),
    (6.5, 3, 1.5, 'PI\nController'),
    (8.5, 3, 1.5, 'Inverse\nPark'),
]

# Draw boxes
for x, y, w, label in boxes:
    box = FancyBboxPatch((x-w/2, y-0.4), w, 0.8,
                         boxstyle="round,pad=0.1",
                         edgecolor='blue', facecolor='lightblue', linewidth=2)
    ax.add_patch(box)
    ax.text(x, y, label, ha='center', va='center', fontsize=10, weight='bold')

# Draw arrows
for i in range(len(boxes)-1):
    x1, y1 = boxes[i][0] + boxes[i][2]/2, boxes[i][1]
    x2, y2 = boxes[i+1][0] - boxes[i+1][2]/2, boxes[i+1][1]
    arrow = FancyArrowPatch((x1, y1), (x2, y2), arrowstyle='->',
                          mutation_scale=20, linewidth=2, color='black')
    ax.add_patch(arrow)

# Add feedback path
feedback_arrow = FancyArrowPatch((9.25, 2.6), (0.5, 2.6),
                               arrowstyle='->', mutation_scale=20,
                               linewidth=1.5, color='red',
                               connectionstyle="arc3,rad=-.5")
ax.add_patch(feedback_arrow)
ax.text(5, 2, 'Feedback Loop', ha='center', fontsize=9, color='red')

# Add PWM stage
pwm_box = FancyBboxPatch((7.8, 4.5), 1.5, 0.8,
                         boxstyle="round,pad=0.1",
                         edgecolor='green', facecolor='lightgreen', linewidth=2)
ax.add_patch(pwm_box)
ax.text(8.55, 4.9, 'PWM\nDriver', ha='center', va='center', fontsize=10, weight='bold')

# Arrow from controller to PWM
arrow_pwm = FancyArrowPatch((8.55, 3.4), (8.55, 4.5),
                           arrowstyle='->', mutation_scale=20, linewidth=2, color='green')
ax.add_patch(arrow_pwm)

# Add motor symbol
motor_box = FancyBboxPatch((8.2, 0.5), 0.7, 0.7,
                          boxstyle="round,pad=0.05",
                          edgecolor='orange', facecolor='lightyellow', linewidth=2)
ax.add_patch(motor_box)
ax.text(8.55, 0.85, 'M', ha='center', va='center', fontsize=14, weight='bold')

# Arrow from PWM to motor
arrow_motor = FancyArrowPatch((8.55, 4.5), (8.55, 1.2),
                             arrowstyle='->', mutation_scale=20, linewidth=1.5, color='orange')
ax.add_patch(arrow_motor)

ax.text(5, 5.5, 'ESP32 FOC Control System Architecture',
        ha='center', fontsize=14, weight='bold')

plt.tight_layout()
plt.savefig('images/foc_architecture.png', dpi=150, bbox_inches='tight')
print("Generated: foc_architecture.png")
plt.close()

# ============================================================
# 2. Three-Phase MOSFET Topology
# ============================================================
fig, ax = plt.subplots(figsize=(10, 8))
ax.set_xlim(-0.5, 10)
ax.set_ylim(-0.5, 8)
ax.axis('off')

def draw_mosfet(ax, x, y, label, orientation='up'):
    """Draw a simple MOSFET symbol"""
    if orientation == 'up':
        # Gate
        ax.plot([x-0.15, x+0.15], [y-0.3, y-0.3], 'k-', linewidth=2)
        # Drain
        ax.plot([x, x], [y+0.4, y+0.8], 'k-', linewidth=2)
        # Source
        ax.plot([x, x], [y-0.8, y-0.4], 'k-', linewidth=2)
        # Body
        rect = mpatches.Rectangle((x-0.2, y-0.15), 0.4, 0.3,
                                 fill=False, edgecolor='black', linewidth=2)
        ax.add_patch(rect)
        ax.text(x+0.5, y, label, fontsize=9, weight='bold')

def draw_resistor(ax, x, y, label, orientation='h'):
    """Draw a simple resistor symbol"""
    if orientation == 'h':
        ax.plot([x-0.5, x-0.2], [y, y], 'k-', linewidth=2)
        rect = mpatches.Rectangle((x-0.2, y-0.1), 0.4, 0.2,
                                 fill=False, edgecolor='black', linewidth=2)
        ax.add_patch(rect)
        ax.plot([x+0.2, x+0.5], [y, y], 'k-', linewidth=2)
    ax.text(x, y+0.4, label, fontsize=8, ha='center')

# Phase A leg
draw_mosfet(ax, 1, 5.5, 'Q1')
draw_mosfet(ax, 1, 2.5, 'Q4')
ax.plot([1, 1], [6.3, 5.7], 'k-', linewidth=2)
ax.plot([1, 1], [1.7, 2.3], 'k-', linewidth=2)
ax.text(0.2, 6.5, '+Vdc', fontsize=10, weight='bold')
ax.plot([0.5, 1.5], [6.8, 6.8], 'k-', linewidth=2)
ax.plot([0.7, 1.3], [1.3, 1.3], 'k-', linewidth=2)
ax.text(0.2, 1.2, 'GND', fontsize=10, weight='bold')
ax.plot([1, 2.5], [4, 4], 'k-', linewidth=1.5)
ax.text(1.5, 4.3, 'Phase A', fontsize=9, weight='bold')

# Phase B leg
draw_mosfet(ax, 4, 5.5, 'Q2')
draw_mosfet(ax, 4, 2.5, 'Q5')
ax.plot([4, 4], [6.3, 5.7], 'k-', linewidth=2)
ax.plot([4, 4], [1.7, 2.3], 'k-', linewidth=2)
ax.plot([3.5, 4.5], [6.8, 6.8], 'k-', linewidth=2)
ax.plot([3.7, 4.3], [1.3, 1.3], 'k-', linewidth=2)
ax.plot([4, 5.5], [4, 4], 'k-', linewidth=1.5)
ax.text(4.5, 4.3, 'Phase B', fontsize=9, weight='bold')

# Phase C leg
draw_mosfet(ax, 7, 5.5, 'Q3')
draw_mosfet(ax, 7, 2.5, 'Q6')
ax.plot([7, 7], [6.3, 5.7], 'k-', linewidth=2)
ax.plot([7, 7], [1.7, 2.3], 'k-', linewidth=2)
ax.plot([6.5, 7.5], [6.8, 6.8], 'k-', linewidth=2)
ax.plot([6.7, 7.3], [1.3, 1.3], 'k-', linewidth=2)
ax.plot([7, 8.5], [4, 4], 'k-', linewidth=1.5)
ax.text(7.5, 4.3, 'Phase C', fontsize=9, weight='bold')

# Motor connection
ax.text(5, 0.3, 'Three-Phase BLDC Motor', fontsize=10, weight='bold', ha='center')
ax.plot([2.5, 4.5], [4, 0.7], 'r-', linewidth=2)
ax.plot([5, 5], [4, 0.7], 'r-', linewidth=2)
ax.plot([5.5, 7.5], [4, 0.7], 'r-', linewidth=2)

ax.text(5, 7.5, 'Three-Phase MOSFET Inverter Topology',
        ha='center', fontsize=12, weight='bold')

plt.tight_layout()
plt.savefig('images/mosfet_topology.png', dpi=150, bbox_inches='tight')
print("Generated: mosfet_topology.png")
plt.close()

# ============================================================
# 3. Clarke and Park Transformations
# ============================================================
fig, axes = plt.subplots(1, 2, figsize=(12, 5))

# Clarke Transform
ax = axes[0]
theta = np.linspace(0, 2*np.pi, 100)
# Three-phase currents
ia = np.sin(theta)
ib = np.sin(theta - 2*np.pi/3)
ic = np.sin(theta - 4*np.pi/3)

ax.plot(theta, ia, 'r-', label='$i_a$', linewidth=2)
ax.plot(theta, ib, 'g-', label='$i_b$', linewidth=2)
ax.plot(theta, ic, 'b-', label='$i_c$', linewidth=2)
ax.set_xlabel('Angle (rad)', fontsize=10)
ax.set_ylabel('Current (A)', fontsize=10)
ax.set_title('Clarke Transform: Three-Phase to Two-Phase', fontsize=11, weight='bold')
ax.legend(loc='upper right')
ax.grid(True, alpha=0.3)
ax.set_xticks([0, np.pi/2, np.pi, 3*np.pi/2, 2*np.pi])
ax.set_xticklabels(['0', '$\pi/2$', '$\pi$', '$3\pi/2$', '$2\pi$'])

# Park Transform
ax = axes[1]
# Park transform output (stationary to rotating reference frame)
omega = 1.0  # electrical angular velocity
t = np.linspace(0, 2*np.pi, 100)
id = 0.5 * np.ones_like(t)  # d-axis component (constant)
iq = 0.7 * np.sin(t)  # q-axis component (oscillating)

ax.plot(t, id, 'r-', label='$i_d$ (constant)', linewidth=2)
ax.plot(t, iq, 'b-', label='$i_q$ (for torque)', linewidth=2)
ax.set_xlabel('Time (s)', fontsize=10)
ax.set_ylabel('Current (A)', fontsize=10)
ax.set_title('Park Transform: Rotating Reference Frame', fontsize=11, weight='bold')
ax.legend(loc='upper right')
ax.grid(True, alpha=0.3)
ax.axhline(y=0, color='k', linestyle='--', alpha=0.3)

plt.tight_layout()
plt.savefig('images/transforms.png', dpi=150, bbox_inches='tight')
print("Generated: transforms.png")
plt.close()

# ============================================================
# 4. Speed and Current Response Performance
# ============================================================
fig, axes = plt.subplots(2, 2, figsize=(12, 9))

# Speed response
ax = axes[0, 0]
t = np.linspace(0, 3, 1000)
speed_ref = 2000 * np.ones_like(t)
speed_actual = speed_ref * (1 - np.exp(-2*t)) * (1 - 0.1*np.exp(-4*t))
ax.plot(t, speed_ref, 'r--', label='Reference', linewidth=2)
ax.plot(t, speed_actual, 'b-', label='Actual', linewidth=2)
ax.set_xlabel('Time (s)', fontsize=10)
ax.set_ylabel('Speed (RPM)', fontsize=10)
ax.set_title('Motor Speed Response', fontsize=11, weight='bold')
ax.legend()
ax.grid(True, alpha=0.3)

# Torque response
ax = axes[0, 1]
torque = 0.8 * (1 - np.exp(-3*t))
ax.plot(t, torque, 'g-', linewidth=2)
ax.set_xlabel('Time (s)', fontsize=10)
ax.set_ylabel('Torque (Nm)', fontsize=10)
ax.set_title('Motor Torque Output', fontsize=11, weight='bold')
ax.grid(True, alpha=0.3)
ax.fill_between(t, 0, torque, alpha=0.3, color='green')

# Current consumption
ax = axes[1, 0]
current = 2.5 * (1 - 0.6*np.exp(-2.5*t)) + 0.3*np.sin(10*t)
ax.plot(t, current, 'purple', linewidth=2)
ax.set_xlabel('Time (s)', fontsize=10)
ax.set_ylabel('Current (A)', fontsize=10)
ax.set_title('Phase Current Consumption', fontsize=11, weight='bold')
ax.grid(True, alpha=0.3)

# Efficiency curve
ax = axes[1, 1]
speeds = np.linspace(500, 3000, 50)
efficiency = 75 + 20*np.exp(-((speeds-2000)**2)/300000)
ax.plot(speeds, efficiency, 'orange', linewidth=2, marker='o', markersize=4)
ax.set_xlabel('Speed (RPM)', fontsize=10)
ax.set_ylabel('Efficiency (%)', fontsize=10)
ax.set_title('Motor Efficiency Curve', fontsize=11, weight='bold')
ax.grid(True, alpha=0.3)
ax.set_ylim([70, 100])

plt.tight_layout()
plt.savefig('images/performance.png', dpi=150, bbox_inches='tight')
print("Generated: performance.png")
plt.close()

# ============================================================
# 5. Control Loop Flow Diagram
# ============================================================
fig, ax = plt.subplots(figsize=(11, 7))
ax.set_xlim(0, 11)
ax.set_ylim(0, 7)
ax.axis('off')

def draw_block(ax, x, y, w, h, label, color='lightblue'):
    """Draw a block with label"""
    box = FancyBboxPatch((x-w/2, y-h/2), w, h,
                         boxstyle="round,pad=0.1",
                         edgecolor='black', facecolor=color, linewidth=2)
    ax.add_patch(box)
    ax.text(x, y, label, ha='center', va='center', fontsize=9, weight='bold')

def draw_arrow(ax, x1, y1, x2, y2, label='', color='black'):
    """Draw an arrow"""
    arrow = FancyArrowPatch((x1, y1), (x2, y2),
                           arrowstyle='->', mutation_scale=15,
                           linewidth=1.5, color=color)
    ax.add_patch(arrow)
    if label:
        mid_x, mid_y = (x1+x2)/2, (y1+y2)/2
        ax.text(mid_x+0.2, mid_y+0.2, label, fontsize=8)

# Draw blocks
draw_block(ax, 1.5, 5.5, 1.2, 0.8, 'Speed Ref.', 'lightyellow')
draw_block(ax, 3.5, 5.5, 1.2, 0.8, 'Speed Error', 'lightcyan')
draw_block(ax, 5.5, 5.5, 1.2, 0.8, 'Speed PI', 'lightgreen')
draw_block(ax, 7.5, 5.5, 1.2, 0.8, 'Id/Iq Ref.', 'lightyellow')

draw_block(ax, 3.5, 3.5, 1.2, 0.8, 'Current Sense', 'lightcoral')
draw_block(ax, 5.5, 3.5, 1.2, 0.8, 'Current Error', 'lightcyan')
draw_block(ax, 7.5, 3.5, 1.2, 0.8, 'Current PI', 'lightgreen')
draw_block(ax, 9.5, 3.5, 1.2, 0.8, 'Inverse Park', 'lightyellow')

draw_block(ax, 7.5, 1.5, 1.2, 0.8, 'SVM PWM', 'lightblue')
draw_block(ax, 9.5, 1.5, 1.2, 0.8, 'MOSFET Driver', 'lightcoral')

# Draw connections
draw_arrow(ax, 2.1, 5.5, 2.9, 5.5)
draw_arrow(ax, 4.1, 5.5, 4.9, 5.5)
draw_arrow(ax, 6.1, 5.5, 6.9, 5.5)
draw_arrow(ax, 7.5, 5.1, 7.5, 3.9)
draw_arrow(ax, 3.5, 3.1, 3.5, 2.3)
draw_arrow(ax, 4.1, 3.5, 4.9, 3.5)
draw_arrow(ax, 6.1, 3.5, 6.9, 3.5)
draw_arrow(ax, 8.1, 3.5, 8.9, 3.5)
draw_arrow(ax, 9.5, 3.1, 9.5, 1.9)
draw_arrow(ax, 8.1, 1.5, 8.9, 1.5)

# Feedback path
draw_arrow(ax, 9.5, 0.9, 9.5, 0.3)
draw_arrow(ax, 9.5, 0.3, 3.5, 0.3)
draw_arrow(ax, 3.5, 0.3, 3.5, 3.1, color='red')

# Motor output
ax.text(10.8, 1.5, 'Three-Phase\nPWM Output', fontsize=9, weight='bold')

ax.text(5.5, 6.5, 'Nested Control Loop Architecture',
        ha='center', fontsize=13, weight='bold')

plt.tight_layout()
plt.savefig('images/control_loop.png', dpi=150, bbox_inches='tight')
print("Generated: control_loop.png")
plt.close()

print("\n✓ All images generated successfully!")
print("Generated images:")
print("  - images/foc_architecture.png")
print("  - images/mosfet_topology.png")
print("  - images/transforms.png")
print("  - images/performance.png")
print("  - images/control_loop.png")
