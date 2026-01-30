import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import numpy as np

# Read data
df = pd.read_excel('/Users/L052239/Downloads/Test/IAA.xlsx')

# Sort by IAA/Receptor value (descending for horizontal bars - highest at top)
df = df.sort_values('IAA/Receptor', ascending=True)

# Create color map (blue to red based on values)
norm = mcolors.Normalize(vmin=df['IAA/Receptor'].min(), vmax=df['IAA/Receptor'].max())
cmap = plt.cm.coolwarm
colors = [cmap(norm(val)) for val in df['IAA/Receptor']]

# Create figure with broken axis (two subplots side by side)
fig, (ax1, ax2) = plt.subplots(1, 2, sharey=True, figsize=(10, 7))
fig.subplots_adjust(wspace=0.05)

# Plot on both axes
ax1.barh(df['TAA'], df['IAA/Receptor'], color=colors)
ax2.barh(df['TAA'], df['IAA/Receptor'], color=colors)

# Set axis limits for the break
ax1.set_xlim(0, 15)
ax2.set_xlim(55, 65)

# Hide spines between the two plots
ax1.spines['right'].set_visible(False)
ax2.spines['left'].set_visible(False)
ax1.yaxis.tick_left()
ax2.yaxis.set_ticks([])

# Ensure y-axis labels are visible
ax1.set_yticks(range(len(df)))
ax1.set_yticklabels(df['TAA'])

# Add diagonal break marks
d = 0.015
kwargs = dict(transform=ax1.transAxes, color='k', clip_on=False)
ax1.plot((1 - d, 1 + d), (-d, +d), **kwargs)
ax1.plot((1 - d, 1 + d), (1 - d, 1 + d), **kwargs)

kwargs.update(transform=ax2.transAxes)
ax2.plot((-d, +d), (-d, +d), **kwargs)
ax2.plot((-d, +d), (1 - d, 1 + d), **kwargs)

# Labels and title
fig.suptitle('IAA/Receptor Values by Gene', fontsize=14, fontweight='bold')
ax1.set_xlabel('IAA/Receptor', fontsize=12)
ax1.xaxis.set_label_coords(1.1, -0.08)


plt.tight_layout()
plt.savefig('/Users/L052239/Downloads/Test/IAA_barplot.png', dpi=300, bbox_inches='tight')
plt.savefig('/Users/L052239/Downloads/Test/IAA_barplot.pdf', bbox_inches='tight')
print("Plot saved as IAA_barplot.png and IAA_barplot.pdf")
