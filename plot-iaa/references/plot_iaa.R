library(readxl)
library(ggplot2)
library(ggbreak)

# Read data
df <- read_excel("/Users/L052239/Downloads/Test/IAA.xlsx")
colnames(df) <- c("TAA", "IAA_Receptor")

# Sort by value and create ordered factor for proper plotting
df <- df[order(df$IAA_Receptor), ]
df$TAA <- factor(df$TAA, levels = df$TAA)

# Create plot with broken x-axis
p <- ggplot(df, aes(x = IAA_Receptor, y = TAA, fill = IAA_Receptor)) +
  geom_bar(stat = "identity") +
  scale_fill_gradientn(colors = c("#3B4CC0", "#7B9FF9", "#C9D7F0", "#F7D4CF", "#F08A6C", "#B40426"),
                       values = scales::rescale(c(1, 10, 20, 30, 40, 60))) +
  scale_x_break(c(15, 55), scales = 0.5) +
  labs(title = "IAA/Receptor Values by Gene",
       x = "IAA/Receptor",
       y = NULL) +
  theme_minimal() +
  theme(legend.position = "none",
        plot.title = element_text(hjust = 0.5, face = "bold", size = 14),
        axis.text.y = element_text(size = 10))

# Save plot
ggsave("/Users/L052239/Downloads/Test/IAA_barplot_R.png", p, width = 10, height = 7, dpi = 300)
ggsave("/Users/L052239/Downloads/Test/IAA_barplot_R.pdf", p, width = 10, height = 7)

print("Plot saved as IAA_barplot_R.png and IAA_barplot_R.pdf")
