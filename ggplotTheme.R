library(ggplot2)

default_theme <- theme_light() +
  theme(
    panel.background = element_blank(),
    panel.border = element_rect(colour = "black"),
    axis.ticks = element_line(colour = "black"),
    axis.text = element_text(size = 12, colour = "black"),
    axis.title = element_text(size = 12),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank(),
    text = element_text(family = "serif", size = 12),
    legend.position = "bottom",
    legend.title = element_text(size = 12, face = "bold"),
    legend.text = element_text(size = 12)
  )
