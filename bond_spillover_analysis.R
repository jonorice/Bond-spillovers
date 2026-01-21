# ============================================================
# Euro area bond yield spillovers: charts + PowerPoint (run-all)
#
# What this script does
# 1) Reads daily 10y yields (nav_test_clean.csv)
# 2) Reads rolling DY spillover output (rolling_EA_all_metrics_*.csv)
# 3) Produces 13 PNG charts with a consistent colour scheme
# 4) Chart07 and Chart08 are 2x2 panels (reliable base plotting)
# 5) Chart13 is robust, uses 2022-2024 and 2024-present, and avoids blank cohorts
# 6) Builds a PPTX in your Downloads directory using officer (no slide_size setter)
# ============================================================

rm(list = ls())

# -----------------------------
# Packages (install if needed)
# -----------------------------
# install.packages(c("zoo","officer"))
library(zoo)
library(officer)

# -----------------------------
# Paths and settings
# -----------------------------
out_dir        <- "C:/Users/ricejon/Downloads"
in_yields_file <- file.path(out_dir, "nav_test_clean.csv")

# Must match your rolling estimation settings
p_use        <- 6
H            <- 25
window       <- 221
table_method <- "avg_exact"

rolling_file <- file.path(
  out_dir,
  paste0("rolling_EA_all_metrics_levels_p", p_use, "_H", H, "_window", window, "_", table_method, ".csv")
)

# Optional PPT template
ppt_template <- file.path(out_dir, "draft slide.pptx")

# Output folders
charts_dir <- file.path(out_dir, "EA_note_charts_FINAL")
data_dir   <- file.path(charts_dir, "chart_data")
if (!dir.exists(charts_dir)) dir.create(charts_dir, recursive = TRUE)
if (!dir.exists(data_dir))   dir.create(data_dir, recursive = TRUE)

# Groups
outside_names <- c("US","JP","UK")
ea_names      <- c("DE","FR","IT","ES")
ea_n          <- length(ea_names)

# Chart date windows
recent_start <- as.Date("2014-01-01")
long_start   <- as.Date("1995-01-01")

# Smoothing windows (trading days)
k_yields_long  <- 63
k_yields_short <- 21
k_spill_smooth <- 63

# -----------------------------
# Prescribed colour scheme (fixed per country throughout)
# -----------------------------
country_cols <- c(
  US = "#003299",
  JP = "#FFB400",
  UK = "#FF4B00",
  DE = "#65B800",
  FR = "#00B1EA",
  IT = "#007816",
  ES = "#8139C6"
)

# Non-country aggregates
col_outside <- "#222222"
col_within  <- "#777777"

# -----------------------------
# Events: numbered markers + keyed labels in footers and slides
# -----------------------------
events <- data.frame(
  id = c(1,  2,  3,  4,  5,  6,  7,  8,  9, 10, 11, 12, 13, 14),
  date = as.Date(c(
    "2012-07-26", # Draghi
    "2015-01-22", # ECB QE announced
    "2016-06-23", # Brexit referendum
    "2016-09-21", # BoJ introduces YCC
    "2018-05-29", # Italy political crisis, BTP sell-off
    "2019-09-12", # ECB restarts QE / tiering
    "2020-03-18", # ECB PEPP
    "2022-02-24", # Ukraine invasion
    "2022-03-16", # Fed hiking cycle begins
    "2022-07-21", # ECB lift-off + TPI announced
    "2022-09-23", # UK mini-budget / gilt crisis
    "2022-12-20", # BoJ widens YCC band
    "2024-03-19", # BoJ ends NIRP / YCC
    "2024-06-06"  # ECB begins cutting cycle (2024)
  )),
  label = c(
    "ECB Draghi "whatever it takes"",
    "ECB announces QE (PSPP)",
    "UK Brexit referendum",
    "BoJ introduces YCC",
    "Italy political crisis, BTP sell-off",
    "ECB restarts QE and tiering",
    "ECB launches PEPP",
    "Russia invades Ukraine",
    "Fed hiking cycle begins",
    "ECB lift-off and TPI",
    "UK mini-budget, gilt crisis",
    "BoJ widens YCC band",
    "BoJ ends NIRP and YCC",
    "ECB begins cutting cycle"
  ),
  stringsAsFactors = FALSE
)

write.csv(events, file.path(data_dir, "events_key.csv"), row.names = FALSE)

# ============================================================
# Helpers
# ============================================================

to_num <- function(x) suppressWarnings(as.numeric(as.character(x)))

# Robust date parsing for common formats and numeric date encodings
parse_date_flexible <- function(x) {
  if (inherits(x, "Date")) return(x)
  
  # Numeric date encodings
  if (is.numeric(x)) {
    # Heuristic:
    # - Excel serial dates typically around 20000-50000 for modern dates (origin 1899-12-30)
    # - Unix days since 1970 typically < 30000 for dates before ~2052
    med <- suppressWarnings(median(x, na.rm = TRUE))
    if (is.finite(med) && med > 30000) {
      return(as.Date(x, origin = "1899-12-30"))
    } else {
      return(as.Date(x, origin = "1970-01-01"))
    }
  }
  
  xc <- trimws(as.character(x))
  d <- suppressWarnings(as.Date(xc))  # ISO first (YYYY-MM-DD)
  
  # Try dd/mm/yyyy
  idx <- is.na(d) & grepl("/", xc)
  if (any(idx)) {
    d[idx] <- suppressWarnings(as.Date(xc[idx], format = "%d/%m/%Y"))
  }
  
  # Try mm/dd/yyyy
  idx2 <- is.na(d) & grepl("/", xc)
  if (any(idx2)) {
    d[idx2] <- suppressWarnings(as.Date(xc[idx2], format = "%m/%d/%Y"))
  }
  
  d
}

rollmean_partial <- function(x, k) {
  zoo::rollapply(
    x, width = k,
    FUN = function(z) mean(z, na.rm = TRUE),
    align = "center", partial = TRUE, fill = NA
  )
}

wrap_lines <- function(lines, width = 170) {
  paste(unlist(lapply(lines, function(s) strwrap(s, width = width))), collapse = "\n")
}

calc_ylim <- function(..., lower0 = FALSE, pad = 0.22) {
  v <- c(...)
  v <- v[is.finite(v)]
  if (length(v) == 0) return(c(0, 1))
  r <- range(v, na.rm = TRUE)
  if (!is.finite(r[1]) || !is.finite(r[2])) return(c(0, 1))
  if (r[1] == r[2]) r <- r + c(-1, 1)
  p <- pad * (r[2] - r[1])
  lo <- r[1] - p
  hi <- r[2] + p
  if (lower0) lo <- 0
  c(lo, hi)
}

calc_ylim_sym <- function(..., pad = 0.25) {
  v <- c(...)
  v <- v[is.finite(v)]
  if (length(v) == 0) return(c(-1, 1))
  m <- max(abs(v), na.rm = TRUE)
  if (!is.finite(m) || m == 0) m <- 1
  m <- m * (1 + pad)
  c(-m, m)
}

filter_events <- function(date_min, date_max, include_ids = NULL) {
  ev <- events[events$date >= date_min & events$date <= date_max, , drop = FALSE]
  if (!is.null(include_ids)) ev <- ev[ev$id %in% include_ids, , drop = FALSE]
  ev
}

event_key_text <- function(ev) {
  if (nrow(ev) == 0) return("Events: none shown.")
  paste0("Events: ", paste0(ev$id, "=", ev$label, collapse = "; "), ".")
}

add_events <- function(ev, ylim, col_line = "#B3B3B3") {
  if (nrow(ev) == 0) return()
  abline(v = ev$date, col = col_line, lty = 3, lwd = 1)
  
  y_top <- ylim[2] - 0.05 * (ylim[2] - ylim[1])
  stagger <- rep(c(0, -0.03, -0.06), length.out = nrow(ev))
  y_pos <- y_top + stagger * (ylim[2] - ylim[1])
  
  text(ev$date, y_pos, labels = ev$id, cex = 0.80, col = "#666666")
}

plot_with_footer <- function(out_png, main_plot_fn, footer_lines,
                             width = 1800, height = 1100, res = 150) {
  png(out_png, width = width, height = height, res = res)
  
  layout(matrix(c(1, 2), nrow = 2), heights = c(4.2, 1))
  par(las = 1)
  par(xpd = FALSE)  # keep drawing inside plot region
  
  par(mar = c(4.2, 4.8, 3.0, 2.0))
  main_plot_fn()
  
  par(mar = c(0.6, 4.8, 0.4, 2.0))
  plot.new()
  footer_text <- wrap_lines(footer_lines, width = 170)
  text(0, 1, labels = footer_text, adj = c(0, 1), cex = 0.95)
  
  dev.off()
}

save_2x2_with_footer_png <- function(out_png, panel_fns, footer_lines,
                                     width = 2000, height = 1250, res = 150) {
  if (length(panel_fns) != 4) stop("panel_fns must be a list of 4 functions.")
  
  png(out_png, width = width, height = height, res = res)
  
  layout(matrix(c(1,2,
                  3,4,
                  5,5), nrow = 3, byrow = TRUE),
         heights = c(1, 1, 0.55))
  par(las = 1)
  par(xpd = FALSE)
  
  for (i in 1:4) {
    par(mar = c(3.6, 4.8, 2.4, 1.4))
    panel_fns[[i]]()
  }
  
  par(mar = c(0.6, 4.8, 0.4, 1.4))
  plot.new()
  footer_text <- wrap_lines(footer_lines, width = 170)
  text(0, 1, labels = footer_text, adj = c(0, 1), cex = 0.95)
  
  dev.off()
}

zscore <- function(x) (x - mean(x, na.rm = TRUE)) / sd(x, na.rm = TRUE)

# ============================================================
# Load yields
# ============================================================
if (!file.exists(in_yields_file)) stop(paste("Yields file not found:", in_yields_file))
y_raw <- read.csv(in_yields_file, stringsAsFactors = FALSE)

y_raw[, 2] <- parse_date_flexible(y_raw[, 2])
if (any(is.na(y_raw[, 2]))) {
  n_bad <- sum(is.na(y_raw[, 2]))
  warning(paste0("Yields file contains ", n_bad, " rows with unparsed dates. These rows will be dropped."))
  y_raw <- y_raw[!is.na(y_raw[, 2]), , drop = FALSE]
}

# Ensure sorted and unique by Date (keep last observation if duplicates)
y_raw <- y_raw[order(y_raw[, 2]), , drop = FALSE]
dup <- duplicated(y_raw[, 2])
if (any(dup)) {
  warning(paste0("Yields file has ", sum(dup), " duplicate dates. Keeping the last observation for each date."))
  y_raw <- y_raw[!duplicated(y_raw[, 2], fromLast = TRUE), , drop = FALSE]
}

yields <- data.frame(Date = y_raw[, 2], y_raw[, -(1:2), drop = FALSE])
for (nm in names(yields)[-1]) yields[[nm]] <- to_num(yields[[nm]])

if (anyNA(yields[-1])) stop("Missing values in yields file. Please re-check interpolation and cleaning.")

need_names <- c(outside_names, ea_names)
if (length(setdiff(need_names, names(yields))) > 0) {
  stop("Column names in yields file do not match expected country codes.")
}

# Derived spreads (basis points) and dispersion
yields$IT_DE_spread_bp <- 100 * (yields$IT - yields$DE)
yields$ES_DE_spread_bp <- 100 * (yields$ES - yields$DE)
yields$FR_DE_spread_bp <- 100 * (yields$FR - yields$DE)
yields$EA_dispersion_sd <- apply(yields[, ea_names, drop = FALSE], 1, sd)

# Smoothed yields
for (nm in c(outside_names, ea_names)) {
  yields[[paste0(nm, "_sm63")]] <- rollmean_partial(yields[[nm]], k_yields_long)
  yields[[paste0(nm, "_sm21")]] <- rollmean_partial(yields[[nm]], k_yields_short)
}

yields$IT_DE_spread_bp_sm63 <- rollmean_partial(yields$IT_DE_spread_bp, k_yields_long)
yields$ES_DE_spread_bp_sm63 <- rollmean_partial(yields$ES_DE_spread_bp, k_yields_long)
yields$FR_DE_spread_bp_sm63 <- rollmean_partial(yields$FR_DE_spread_bp, k_yields_long)

yields$IT_DE_spread_bp_sm21 <- rollmean_partial(yields$IT_DE_spread_bp, k_yields_short)
yields$EA_dispersion_sd_sm21 <- rollmean_partial(yields$EA_dispersion_sd, k_yields_short)

# ============================================================
# Load rolling spillover metrics
# ============================================================
if (!file.exists(rolling_file)) stop(paste("Rolling metrics file not found:", rolling_file))

roll <- read.csv(rolling_file, stringsAsFactors = FALSE)
roll$Date <- parse_date_flexible(roll$Date)

if (any(is.na(roll$Date))) {
  n_bad <- sum(is.na(roll$Date))
  warning(paste0("Rolling file contains ", n_bad, " rows with unparsed dates. These rows will be dropped."))
  roll <- roll[!is.na(roll$Date), , drop = FALSE]
}

roll <- roll[order(roll$Date), , drop = FALSE]

# Create/repair core origin series if needed
ensure_origin_avg_pct <- function(origin) {
  c_avg <- paste0("EA_from_", origin, "_avg_pct")
  c_pp  <- paste0("EA_from_", origin, "_pp")
  
  if (!(c_avg %in% names(roll)) && (c_pp %in% names(roll))) {
    roll[[c_avg]] <- to_num(roll[[c_pp]]) / ea_n
  }
  if (!(c_pp %in% names(roll)) && (c_avg %in% names(roll))) {
    roll[[c_pp]] <- to_num(roll[[c_avg]]) * ea_n
  }
  
  invisible(TRUE)
}

for (o in c(outside_names, ea_names)) ensure_origin_avg_pct(o)

# Ensure EA_from_outside_avg_pct and EA_within_avg_pct exist
if (!("EA_from_outside_avg_pct" %in% names(roll))) {
  roll$EA_from_outside_avg_pct <- to_num(roll$EA_from_US_avg_pct) +
    to_num(roll$EA_from_JP_avg_pct) + to_num(roll$EA_from_UK_avg_pct)
}
if (!("EA_within_avg_pct" %in% names(roll))) {
  roll$EA_within_avg_pct <- to_num(roll$EA_from_DE_avg_pct) + to_num(roll$EA_from_FR_avg_pct) +
    to_num(roll$EA_from_IT_avg_pct) + to_num(roll$EA_from_ES_avg_pct)
}

# Smooth spillover series used in charts
need_cols <- c(
  "EA_from_outside_avg_pct", "EA_within_avg_pct",
  "EA_from_US_avg_pct", "EA_from_JP_avg_pct", "EA_from_UK_avg_pct",
  "EA_from_DE_avg_pct", "EA_from_FR_avg_pct", "EA_from_IT_avg_pct", "EA_from_ES_avg_pct"
)
miss <- setdiff(need_cols, names(roll))
if (length(miss) > 0) stop(paste("Missing required columns in rolling file:", paste(miss, collapse = ", ")))

for (nm in need_cols) {
  roll[[paste0(nm, "_sm")]] <- rollmean_partial(to_num(roll[[nm]]), k_spill_smooth)
}

# Net within-EA transmitter measure from bilaterals (within EA only)
incoming_from_ea_pp <- function(receiver) {
  origins <- setdiff(ea_names, receiver)
  cols <- paste0(origins, "_to_", receiver)
  miss_cols <- setdiff(cols, names(roll))
  if (length(miss_cols) > 0) stop(paste0("Missing bilateral columns for receiver ", receiver, ": ", paste(miss_cols, collapse = ", ")))
  rowSums(roll[, cols, drop = FALSE])
}

k_links <- ea_n - 1

to_pp <- list(
  DE = to_num(roll$EA_from_DE_avg_pct) * ea_n,
  FR = to_num(roll$EA_from_FR_avg_pct) * ea_n,
  IT = to_num(roll$EA_from_IT_avg_pct) * ea_n,
  ES = to_num(roll$EA_from_ES_avg_pct) * ea_n
)

from_pp <- list(
  DE = incoming_from_ea_pp("DE"),
  FR = incoming_from_ea_pp("FR"),
  IT = incoming_from_ea_pp("IT"),
  ES = incoming_from_ea_pp("ES")
)

roll$DE_net_withinEA <- (to_pp$DE - from_pp$DE) / k_links
roll$FR_net_withinEA <- (to_pp$FR - from_pp$FR) / k_links
roll$IT_net_withinEA <- (to_pp$IT - from_pp$IT) / k_links
roll$ES_net_withinEA <- (to_pp$ES - from_pp$ES) / k_links

for (nm in c("DE_net_withinEA","FR_net_withinEA","IT_net_withinEA","ES_net_withinEA")) {
  roll[[paste0(nm, "_sm")]] <- rollmean_partial(to_num(roll[[nm]]), k_spill_smooth)
}

# ============================================================
# CHART 01 Long run yields (US, UK, JP, DE, IT)
# ============================================================
chart01 <- yields[yields$Date >= long_start, c("Date", "US_sm63","UK_sm63","JP_sm63","DE_sm63","IT_sm63")]
write.csv(chart01, file.path(data_dir, "chart01_longrun_yields_sm63.csv"), row.names = FALSE)

plot_with_footer(
  file.path(charts_dir, "chart01_longrun_yields_sm63.png"),
  main_plot_fn = function() {
    y_lim <- calc_ylim(chart01$US_sm63, chart01$UK_sm63, chart01$JP_sm63, chart01$DE_sm63, chart01$IT_sm63, pad = 0.22)
    plot(chart01$Date, chart01$US_sm63, type = "l", lwd = 2.8, col = country_cols["US"],
         xlab = "Date", ylab = "10 year yield (percent)",
         main = "Long run 10 year yields (smoothed)", ylim = y_lim)
    lines(chart01$Date, chart01$UK_sm63, lwd = 2.8, col = country_cols["UK"])
    lines(chart01$Date, chart01$JP_sm63, lwd = 2.8, col = country_cols["JP"])
    lines(chart01$Date, chart01$DE_sm63, lwd = 2.8, col = country_cols["DE"])
    lines(chart01$Date, chart01$IT_sm63, lwd = 2.8, col = country_cols["IT"])
    graphics::grid(col = "#E6E6E6", lty = 1)
    ev <- filter_events(min(chart01$Date, na.rm=TRUE), max(chart01$Date, na.rm=TRUE),
                        include_ids = c(1,2,3,4,5,6,7,8,9,10,11,12,13,14))
    add_events(ev, y_lim)
    legend("topright",
           legend = c("US","UK","JP","DE","IT"),
           col = country_cols[c("US","UK","JP","DE","IT")],
           lty = 1, lwd = 2.8, bty = "n")
  },
  footer_lines = c(
    "What it shows, long-run rate regimes across major sovereign markets.",
    paste0("Method, daily 10 year yields smoothed using a ", k_yields_long, "-day centred moving average (partial windows at the ends)."),
    event_key_text(filter_events(min(chart01$Date, na.rm=TRUE), max(chart01$Date, na.rm=TRUE),
                                 include_ids = c(2,3,4,5,6,7,8,9,10,11,12,13,14)))
  )
)

# ============================================================
# CHART 02 Long run EA spreads to Germany (basis points)
# ============================================================
chart02 <- yields[yields$Date >= long_start, c("Date","IT_DE_spread_bp_sm63","ES_DE_spread_bp_sm63","FR_DE_spread_bp_sm63")]
write.csv(chart02, file.path(data_dir, "chart02_EA_spreads_to_DE_bp_sm63.csv"), row.names = FALSE)

plot_with_footer(
  file.path(charts_dir, "chart02_EA_spreads_to_DE_bp_sm63.png"),
  main_plot_fn = function() {
    y_lim <- calc_ylim(chart02$IT_DE_spread_bp_sm63, chart02$ES_DE_spread_bp_sm63, chart02$FR_DE_spread_bp_sm63, lower0 = TRUE, pad = 0.22)
    plot(chart02$Date, chart02$IT_DE_spread_bp_sm63, type = "l", lwd = 2.8, col = country_cols["IT"],
         xlab = "Date", ylab = "Spread to Germany (basis points)",
         main = "Euro area spreads to Germany (smoothed)", ylim = y_lim)
    lines(chart02$Date, chart02$ES_DE_spread_bp_sm63, lwd = 2.8, col = country_cols["ES"])
    lines(chart02$Date, chart02$FR_DE_spread_bp_sm63, lwd = 2.8, col = country_cols["FR"])
    graphics::grid(col = "#E6E6E6", lty = 1)
    ev <- filter_events(min(chart02$Date, na.rm=TRUE), max(chart02$Date, na.rm=TRUE),
                        include_ids = c(2,5,6,7,8,10,11,14))
    add_events(ev, y_lim)
    legend("topright",
           legend = c("IT minus DE","ES minus DE","FR minus DE"),
           col = country_cols[c("IT","ES","FR")],
           lty = 1, lwd = 2.8, bty = "n")
  },
  footer_lines = c(
    "What it shows, EA fragmentation episodes through periphery and France spreads to Germany.",
    paste0("Method, spreads computed from daily 10 year yields (bp) and smoothed with a ", k_yields_long, "-day centred moving average."),
    event_key_text(filter_events(min(chart02$Date, na.rm=TRUE), max(chart02$Date, na.rm=TRUE),
                                 include_ids = c(2,5,6,7,8,10,11,14)))
  )
)

# ============================================================
# CHART 03 Recent yields: DE and IT versus US, UK, JP
# ============================================================
chart03 <- yields[yields$Date >= recent_start, c("Date","DE_sm21","IT_sm21","US_sm21","UK_sm21","JP_sm21")]
write.csv(chart03, file.path(data_dir, "chart03_recent_yields_DE_IT_outside_sm21.csv"), row.names = FALSE)

plot_with_footer(
  file.path(charts_dir, "chart03_recent_yields_DE_IT_outside_sm21.png"),
  main_plot_fn = function() {
    y_lim <- calc_ylim(chart03$DE_sm21, chart03$IT_sm21, chart03$US_sm21, chart03$UK_sm21, chart03$JP_sm21, pad = 0.22)
    plot(chart03$Date, chart03$DE_sm21, type = "l", lwd = 2.8, col = country_cols["DE"],
         xlab = "Date", ylab = "10 year yield (percent)",
         main = "Recent yields, Germany and Italy versus US, UK, Japan (smoothed)", ylim = y_lim)
    lines(chart03$Date, chart03$IT_sm21, lwd = 2.8, col = country_cols["IT"])
    lines(chart03$Date, chart03$US_sm21, lwd = 2.8, col = country_cols["US"])
    lines(chart03$Date, chart03$UK_sm21, lwd = 2.8, col = country_cols["UK"])
    lines(chart03$Date, chart03$JP_sm21, lwd = 2.8, col = country_cols["JP"])
    graphics::grid(col = "#E6E6E6", lty = 1)
    ev <- filter_events(min(chart03$Date, na.rm=TRUE), max(chart03$Date, na.rm=TRUE),
                        include_ids = c(2,3,4,5,6,7,8,9,10,11,12,13,14))
    add_events(ev, y_lim)
    legend("topright",
           legend = c("DE","IT","US","UK","JP"),
           col = country_cols[c("DE","IT","US","UK","JP")],
           lty = 1, lwd = 2.8, bty = "n")
  },
  footer_lines = c(
    "What it shows, post-2014 rate dynamics with DE representing EA core and IT representing EA periphery.",
    paste0("Method, daily 10 year yields smoothed with a ", k_yields_short, "-day centred moving average."),
    event_key_text(filter_events(min(chart03$Date, na.rm=TRUE), max(chart03$Date, na.rm=TRUE),
                                 include_ids = c(2,3,4,5,6,7,8,9,10,11,12,13,14)))
  )
)

# ============================================================
# CHART 04 Fragmentation vs dispersion (standardised)
# ============================================================
df4 <- yields[yields$Date >= recent_start, c("Date","IT_DE_spread_bp_sm21","EA_dispersion_sd_sm21")]
df4$IT_DE_z   <- zscore(df4$IT_DE_spread_bp_sm21)
df4$EA_disp_z <- zscore(df4$EA_dispersion_sd_sm21)
chart04 <- df4[, c("Date","IT_DE_z","EA_disp_z")]
write.csv(chart04, file.path(data_dir, "chart04_fragmentation_vs_dispersion_z.csv"), row.names = FALSE)

plot_with_footer(
  file.path(charts_dir, "chart04_fragmentation_vs_dispersion_z.png"),
  main_plot_fn = function() {
    y_lim <- calc_ylim_sym(chart04$IT_DE_z, chart04$EA_disp_z, pad = 0.25)
    plot(chart04$Date, chart04$IT_DE_z, type = "l", lwd = 2.8, col = country_cols["IT"],
         xlab = "Date", ylab = "Standardised index",
         main = "EA fragmentation and cross-country dispersion (standardised)", ylim = y_lim)
    lines(chart04$Date, chart04$EA_disp_z, lwd = 2.8, col = "#444444")
    abline(h = 0, lty = 3, col = "#888888")
    graphics::grid(col = "#E6E6E6", lty = 1)
    ev <- filter_events(min(chart04$Date, na.rm=TRUE), max(chart04$Date, na.rm=TRUE),
                        include_ids = c(2,5,6,7,8,10,11,14))
    add_events(ev, y_lim)
    legend("topright",
           legend = c("IT-DE spread (z)","EA dispersion (z)"),
           col = c(country_cols["IT"], "#444444"),
           lty = 1, lwd = 2.8, bty = "n")
  },
  footer_lines = c(
    "What it shows, whether EA dispersion co-moves with fragmentation in Italy spreads.",
    paste0("Method, IT-DE (bp) and EA cross-sectional SD are smoothed (", k_yields_short, " days) then standardised."),
    event_key_text(filter_events(min(chart04$Date, na.rm=TRUE), max(chart04$Date, na.rm=TRUE),
                                 include_ids = c(2,5,6,7,8,10,11,14)))
  )
)

# ============================================================
# CHART 05 Spillovers into EA: outside vs within (avg per EA member)
# ============================================================
chart05 <- roll[roll$Date >= recent_start, c("Date","EA_from_outside_avg_pct_sm","EA_within_avg_pct_sm")]
write.csv(chart05, file.path(data_dir, "chart05_spill_into_EA_outside_vs_within.csv"), row.names = FALSE)

plot_with_footer(
  file.path(charts_dir, "chart05_spill_into_EA_outside_vs_within.png"),
  main_plot_fn = function() {
    y_lim <- calc_ylim(chart05$EA_from_outside_avg_pct_sm, chart05$EA_within_avg_pct_sm, lower0 = TRUE, pad = 0.25)
    plot(chart05$Date, chart05$EA_from_outside_avg_pct_sm, type = "l", lwd = 3.0, col = col_outside,
         xlab = "Date", ylab = "Average spillover into an EA member (percent)",
         main = "Rolling spillovers into EA, outside versus within (smoothed)", ylim = y_lim)
    lines(chart05$Date, chart05$EA_within_avg_pct_sm, lwd = 3.0, col = col_within)
    graphics::grid(col = "#E6E6E6", lty = 1)
    ev <- filter_events(min(chart05$Date, na.rm=TRUE), max(chart05$Date, na.rm=TRUE),
                        include_ids = c(7,8,9,10,11,12,13,14))
    add_events(ev, y_lim)
    legend("topright",
           legend = c("From outside (US, JP, UK)","Within EA (DE, FR, IT, ES)"),
           col = c(col_outside, col_within),
           lty = 1, lwd = 3.0, bty = "n")
  },
  footer_lines = c(
    "What it shows, whether EA yield variance is dominated by external or internal shocks.",
    paste0("Method, rolling VAR(", p_use, ") FEVD spillovers, window ", window, ", horizon ", H, ", averaged across DE/FR/IT/ES and smoothed (", k_spill_smooth, " days)."),
    event_key_text(filter_events(min(chart05$Date, na.rm=TRUE), max(chart05$Date, na.rm=TRUE),
                                 include_ids = c(7,8,9,10,11,12,13,14)))
  )
)

# ============================================================
# CHART 06 External sources into EA (US, JP, UK)
# ============================================================
chart06 <- roll[roll$Date >= recent_start, c("Date","EA_from_US_avg_pct_sm","EA_from_JP_avg_pct_sm","EA_from_UK_avg_pct_sm")]
write.csv(chart06, file.path(data_dir, "chart06_external_sources_into_EA.csv"), row.names = FALSE)

plot_with_footer(
  file.path(charts_dir, "chart06_external_sources_into_EA.png"),
  main_plot_fn = function() {
    y_lim <- calc_ylim(chart06$EA_from_US_avg_pct_sm, chart06$EA_from_JP_avg_pct_sm, chart06$EA_from_UK_avg_pct_sm, lower0 = TRUE, pad = 0.25)
    plot(chart06$Date, chart06$EA_from_US_avg_pct_sm, type = "l", lwd = 3.0, col = country_cols["US"],
         xlab = "Date", ylab = "Average spillover into an EA member (percent)",
         main = "External spillovers into EA by source (smoothed)", ylim = y_lim)
    lines(chart06$Date, chart06$EA_from_JP_avg_pct_sm, lwd = 3.0, col = country_cols["JP"])
    lines(chart06$Date, chart06$EA_from_UK_avg_pct_sm, lwd = 3.0, col = country_cols["UK"])
    graphics::grid(col = "#E6E6E6", lty = 1)
    ev <- filter_events(min(chart06$Date, na.rm=TRUE), max(chart06$Date, na.rm=TRUE),
                        include_ids = c(7,8,9,11,12,13,14))
    add_events(ev, y_lim)
    legend("topright",
           legend = c("US","JP","UK"),
           col = country_cols[c("US","JP","UK")],
           lty = 1, lwd = 3.0, bty = "n")
  },
  footer_lines = c(
    "What it shows, which external market is most informative for EA yields in variance-decomposition terms.",
    paste0("Method, contributions into EA are averaged across DE/FR/IT/ES and smoothed (", k_spill_smooth, " days)."),
    event_key_text(filter_events(min(chart06$Date, na.rm=TRUE), max(chart06$Date, na.rm=TRUE),
                                 include_ids = c(7,8,9,11,12,13,14)))
  )
)

# ============================================================
# CHART 07 Within-EA sources into EA (2x2 panels)
# ============================================================
df7 <- roll[roll$Date >= recent_start,
            c("Date","EA_from_DE_avg_pct","EA_from_FR_avg_pct","EA_from_IT_avg_pct","EA_from_ES_avg_pct")]
for (cc in names(df7)[-1]) df7[[cc]] <- rollmean_partial(to_num(df7[[cc]]), k_spill_smooth)
write.csv(df7, file.path(data_dir, "chart07_internal_sources_into_EA.csv"), row.names = FALSE)

y_lim_7 <- calc_ylim(df7$EA_from_DE_avg_pct, df7$EA_from_FR_avg_pct, df7$EA_from_IT_avg_pct, df7$EA_from_ES_avg_pct,
                     lower0 = TRUE, pad = 0.25)
ev7 <- filter_events(min(df7$Date, na.rm=TRUE), max(df7$Date, na.rm=TRUE), include_ids = c(7,8,10,11,12,13,14))

panel_fns_7 <- list(
  function() {
    plot(df7$Date, df7$EA_from_DE_avg_pct, type="l", lwd=2.8, col=country_cols["DE"], ylim=y_lim_7,
         xlab="Date", ylab="Avg spillover into EA member (percent)", main="Origin: DE")
    graphics::grid(col="#E6E6E6", lty=1); add_events(ev7, y_lim_7)
  },
  function() {
    plot(df7$Date, df7$EA_from_FR_avg_pct, type="l", lwd=2.8, col=country_cols["FR"], ylim=y_lim_7,
         xlab="Date", ylab="Avg spillover into EA member (percent)", main="Origin: FR")
    graphics::grid(col="#E6E6E6", lty=1); add_events(ev7, y_lim_7)
  },
  function() {
    plot(df7$Date, df7$EA_from_IT_avg_pct, type="l", lwd=2.8, col=country_cols["IT"], ylim=y_lim_7,
         xlab="Date", ylab="Avg spillover into EA member (percent)", main="Origin: IT")
    graphics::grid(col="#E6E6E6", lty=1); add_events(ev7, y_lim_7)
  },
  function() {
    plot(df7$Date, df7$EA_from_ES_avg_pct, type="l", lwd=2.8, col=country_cols["ES"], ylim=y_lim_7,
         xlab="Date", ylab="Avg spillover into EA member (percent)", main="Origin: ES")
    graphics::grid(col="#E6E6E6", lty=1); add_events(ev7, y_lim_7)
  }
)

save_2x2_with_footer_png(
  file.path(charts_dir, "chart07_internal_sources_into_EA_2x2.png"),
  panel_fns_7,
  footer_lines = c(
    "What it shows, within-EA contributions to spillovers into the EA block, split by origin.",
    paste0("Method, rolling VAR(", p_use, ") FEVD spillovers, averaged across receivers (DE, FR, IT, ES), smoothed (", k_spill_smooth, " days)."),
    event_key_text(ev7)
  )
)

# ============================================================
# CHART 08 Net within-EA transmitters (2x2 panels)
# ============================================================
df8 <- roll[roll$Date >= recent_start,
            c("Date","DE_net_withinEA_sm","FR_net_withinEA_sm","IT_net_withinEA_sm","ES_net_withinEA_sm")]
write.csv(df8, file.path(data_dir, "chart08_net_withinEA.csv"), row.names = FALSE)

y_lim_8 <- calc_ylim_sym(df8$DE_net_withinEA_sm, df8$FR_net_withinEA_sm, df8$IT_net_withinEA_sm, df8$ES_net_withinEA_sm, pad = 0.25)
ev8 <- filter_events(min(df8$Date, na.rm=TRUE), max(df8$Date, na.rm=TRUE), include_ids = c(7,8,10,11,12,13,14))

panel_fns_8 <- list(
  function() {
    plot(df8$Date, df8$DE_net_withinEA_sm, type="l", lwd=2.8, col=country_cols["DE"], ylim=y_lim_8,
         xlab="Date", ylab="Net within-EA spillover", main="Net transmitter: DE")
    abline(h=0, lty=3, col="#888888"); graphics::grid(col="#E6E6E6", lty=1); add_events(ev8, y_lim_8)
  },
  function() {
    plot(df8$Date, df8$FR_net_withinEA_sm, type="l", lwd=2.8, col=country_cols["FR"], ylim=y_lim_8,
         xlab="Date", ylab="Net within-EA spillover", main="Net transmitter: FR")
    abline(h=0, lty=3, col="#888888"); graphics::grid(col="#E6E6E6", lty=1); add_events(ev8, y_lim_8)
  },
  function() {
    plot(df8$Date, df8$IT_net_withinEA_sm, type="l", lwd=2.8, col=country_cols["IT"], ylim=y_lim_8,
         xlab="Date", ylab="Net within-EA spillover", main="Net transmitter: IT")
    abline(h=0, lty=3, col="#888888"); graphics::grid(col="#E6E6E6", lty=1); add_events(ev8, y_lim_8)
  },
  function() {
    plot(df8$Date, df8$ES_net_withinEA_sm, type="l", lwd=2.8, col=country_cols["ES"], ylim=y_lim_8,
         xlab="Date", ylab="Net within-EA spillover", main="Net transmitter: ES")
    abline(h=0, lty=3, col="#888888"); graphics::grid(col="#E6E6E6", lty=1); add_events(ev8, y_lim_8)
  }
)

save_2x2_with_footer_png(
  file.path(charts_dir, "chart08_net_withinEA_2x2.png"),
  panel_fns_8,
  footer_lines = c(
    "What it shows, directionality within EA. Positive means a net transmitter to other EA members.",
    paste0("Method, net = outgoing minus incoming within-EA spillovers (scaled per bilateral link), smoothed (", k_spill_smooth, " days)."),
    event_key_text(ev8)
  )
)

# ============================================================
# CHARTS 09-12 Receiver-specific within-EA incoming spillovers
# ============================================================
make_receiver_chart <- function(receiver) {
  
  origins <- setdiff(ea_names, receiver)
  cols <- paste0(origins, "_to_", receiver)
  
  miss_cols <- setdiff(cols, names(roll))
  if (length(miss_cols) > 0) stop(paste0("Missing bilateral columns for receiver ", receiver, ": ", paste(miss_cols, collapse=", ")))
  
  df <- roll[roll$Date >= recent_start, c("Date", cols), drop = FALSE]
  for (cc in cols) df[[cc]] <- rollmean_partial(to_num(df[[cc]]), k_spill_smooth)
  
  write.csv(df, file.path(data_dir, paste0("chart_incoming_withinEA_to_", receiver, ".csv")), row.names = FALSE)
  
  out_png <- file.path(charts_dir, paste0("chart_incoming_withinEA_to_", receiver, ".png"))
  
  ev <- filter_events(min(df$Date, na.rm=TRUE), max(df$Date, na.rm=TRUE), include_ids = c(7,8,10,11,12,13,14))
  
  plot_with_footer(
    out_png,
    main_plot_fn = function() {
      y_lim <- calc_ylim(df[[cols[1]]], df[[cols[2]]], df[[cols[3]]], lower0 = TRUE, pad = 0.25)
      plot(df$Date, df[[cols[1]]], type="l", lwd=3.0, col=country_cols[origins[1]],
           xlab="Date", ylab="Spillover into receiver (percent)",
           main=paste0("Within-EA incoming spillovers into ", receiver, " (smoothed)"),
           ylim=y_lim)
      lines(df$Date, df[[cols[2]]], lwd=3.0, col=country_cols[origins[2]])
      lines(df$Date, df[[cols[3]]], lwd=3.0, col=country_cols[origins[3]])
      graphics::grid(col="#E6E6E6", lty=1)
      add_events(ev, y_lim)
      legend("topright", legend = cols, col = country_cols[origins], lty=1, lwd=3.0, bty="n")
    },
    footer_lines = c(
      paste0("What it shows, which EA markets explain movements in ", receiver, " in variance-decomposition terms."),
      paste0("Method, bilateral DY spillovers from origin to receiver, smoothed (", k_spill_smooth, " days)."),
      event_key_text(ev)
    )
  )
  
  invisible(TRUE)
}

make_receiver_chart("DE")
make_receiver_chart("FR")
make_receiver_chart("IT")
make_receiver_chart("ES")

# ============================================================
# CHART 13 Regime summary, spillovers into EA by origin (FIXED, robust)
# Regimes: 2015-2019, 2020-2021, 2022-2024, 2024-present (buffered end)
#
# Important, this chart does NOT require external bilaterals (US_to_DE etc).
# It uses EA_from_<origin>_avg_pct (or EA_from_<origin>_pp / ea_n) and will
# automatically choose the best available series to avoid empty cohorts.
# ============================================================

origins_c13 <- c("US","JP","UK","DE","FR","IT","ES")

# Candidate builder per origin
get_c13_candidates <- function(origin) {
  c_avg    <- paste0("EA_from_", origin, "_avg_pct")
  c_avg_sm <- paste0(c_avg, "_sm")
  c_pp     <- paste0("EA_from_", origin, "_pp")
  c_pp_sm  <- paste0(c_pp, "_sm")
  
  out <- list()
  
  if (c_avg %in% names(roll))    out[[c_avg]]    <- to_num(roll[[c_avg]])
  if (c_avg_sm %in% names(roll)) out[[c_avg_sm]] <- to_num(roll[[c_avg_sm]])
  if (c_pp %in% names(roll))     out[[c_pp]]     <- to_num(roll[[c_pp]]) / ea_n
  if (c_pp_sm %in% names(roll))  out[[c_pp_sm]]  <- to_num(roll[[c_pp_sm]]) / ea_n
  
  # Bilateral fallback for EA origins only, if you have those columns
  cols_bi <- paste0(origin, "_to_", ea_names)
  if (all(cols_bi %in% names(roll))) {
    out[["bilateral_mean"]] <- rowMeans(roll[, cols_bi, drop = FALSE], na.rm = TRUE)
  }
  
  if (length(out) == 0) {
    stop(paste0("Chart13 cannot find any usable origin series for ", origin,
                ". Expected EA_from_", origin, "_avg_pct or EA_from_", origin, "_pp."))
  }
  
  out
}

# Select best series per origin, prioritising coverage from 2024 onward, then from 2022 onward
select_best_c13 <- function(origin, date_vec) {
  cand <- get_c13_candidates(origin)
  nm <- names(cand)
  
  d2022 <- as.Date("2022-01-01")
  d2024 <- as.Date("2024-01-01")
  
  pref_order <- c(
    paste0("EA_from_", origin, "_avg_pct"),
    paste0("EA_from_", origin, "_avg_pct_sm"),
    paste0("EA_from_", origin, "_pp"),
    paste0("EA_from_", origin, "_pp_sm"),
    "bilateral_mean"
  )
  
  score <- sapply(nm, function(k) {
    x <- cand[[k]]
    if (length(x) != length(date_vec)) return(-Inf)
    
    n_2024 <- sum(!is.na(x) & date_vec >= d2024)
    n_2022 <- sum(!is.na(x) & date_vec >= d2022)
    last_dt <- suppressWarnings(max(date_vec[!is.na(x)], na.rm = TRUE))
    last_num <- if (is.finite(last_dt)) as.numeric(last_dt) else 0
    
    pref <- match(k, pref_order)
    if (is.na(pref)) pref <- 99
    
    # Heavily weight post-2024 coverage to prevent blank 2024-present bars
    n_2024 * 1e8 + n_2022 * 1e4 + last_num - pref
  })
  
  best <- nm[which.max(score)]
  list(chosen = best, vec = cand[[best]])
}

# Build panel
X13 <- data.frame(Date = roll$Date)
series_used <- data.frame(origin = origins_c13, chosen_series = NA_character_, stringsAsFactors = FALSE)

for (i in seq_along(origins_c13)) {
  o <- origins_c13[i]
  sel <- select_best_c13(o, roll$Date)
  X13[[o]] <- sel$vec
  series_used$chosen_series[i] <- sel$chosen
}

write.csv(series_used, file.path(data_dir, "chart13_series_used.csv"), row.names = FALSE)

# Robust end date selection with buffered end
# Use a buffer at least 5 days, and also at least half the smoothing window to avoid any trailing gaps
buffer_days <- max(5, ceiling(k_spill_smooth / 2))

valid_row <- !is.na(X13$Date) & (rowSums(!is.na(X13[, origins_c13, drop = FALSE])) > 0)
if (!any(valid_row)) stop("Chart13, no usable values found for any origin in the rolling file.")

last_valid_date <- max(X13$Date[valid_row], na.rm = TRUE)
end_present <- last_valid_date - buffer_days
if (!is.finite(end_present) || end_present < min(X13$Date[valid_row], na.rm = TRUE)) end_present <- last_valid_date

# Define regimes as requested
regimes <- data.frame(
  regime = c("2015-2019", "2020-2021", "2022-2024", "2024-present"),
  start  = as.Date(c("2015-01-01", "2020-01-01", "2022-01-01", "2024-01-01")),
  end    = as.Date(c("2019-12-31", "2021-12-31", min(as.Date("2024-12-31"), end_present), end_present)),
  stringsAsFactors = FALSE
)

# Compute regime means
reg_mat <- matrix(NA_real_, nrow = length(origins_c13), ncol = nrow(regimes))
rownames(reg_mat) <- origins_c13
colnames(reg_mat) <- regimes$regime

reg_counts <- integer(nrow(regimes))

for (r in 1:nrow(regimes)) {
  idx <- (X13$Date >= regimes$start[r]) & (X13$Date <= regimes$end[r])
  idx[is.na(idx)] <- FALSE
  
  sub <- X13[idx, origins_c13, drop = FALSE]
  sub <- sub[rowSums(!is.na(sub)) > 0, , drop = FALSE]
  reg_counts[r] <- nrow(sub)
  
  if (nrow(sub) > 0) {
    m <- colMeans(sub, na.rm = TRUE)
    m[is.nan(m)] <- NA_real_
    reg_mat[, r] <- m
  }
}

# Hard checks so you never silently get a blank cohort again
if (any(reg_counts == 0)) {
  stop(paste0(
    "Chart13, one or more regimes contain zero usable observations. Counts ",
    paste0(regimes$regime, "=", reg_counts, collapse = ", "),
    ". This usually means your rolling file does not contain data through that period, or Date parsing failed."
  ))
}

col_all_na <- apply(reg_mat, 2, function(z) all(is.na(z)))
if (any(col_all_na)) {
  stop(paste0(
    "Chart13, regime means are all NA for ",
    paste(names(which(col_all_na)), collapse = ", "),
    ". Inspect chart13_series_used.csv and confirm EA_from_<origin> series have data through 2024 and beyond."
  ))
}

write.csv(
  cbind(origin = rownames(reg_mat), reg_mat),
  file.path(data_dir, "chart13_regime_summary_FIXED.csv"),
  row.names = FALSE
)

chart13_png <- file.path(charts_dir, "chart13_regime_summary_FIXED.png")

plot_with_footer(
  chart13_png,
  main_plot_fn = function() {
    y_top <- max(reg_mat, na.rm = TRUE)
    if (!is.finite(y_top) || y_top <= 0) y_top <- 1
    y_lim <- c(0, y_top * 1.25)
    
    bp <- barplot(reg_mat,
                  beside = TRUE,
                  col = country_cols[rownames(reg_mat)],
                  border = "black",
                  ylim = y_lim,
                  main = "Regime summary, spillovers into EA by origin",
                  xlab = "Regime",
                  ylab = "Average spillover into an EA member (percent)")
    
    legend("topright", legend = rownames(reg_mat),
           fill = country_cols[rownames(reg_mat)], bty = "n")
  },
  footer_lines = c(
    "What it shows, how the average importance of each origin changes across regimes.",
    paste0("Method, regime means of spillovers into DE, FR, IT, ES by origin. 2024-present ends ", buffer_days, " days before the last usable observation."),
    paste0("Regimes, 2015-2019, 2020-2021, 2022-2024, 2024-present (to ", format(end_present), ").")
  ),
  width = 1800, height = 1050, res = 150
)

# ============================================================
# Build PowerPoint (template-driven, no slide_size setter)
# ============================================================

ppt_file <- file.path(out_dir, "EA_bond_spillovers_deck_FINAL.pptx")

# Read from template if it exists, else default
if (nzchar(ppt_template) && file.exists(ppt_template)) {
  ppt <- read_pptx(path = ppt_template)
} else {
  ppt <- read_pptx()
}

# Slide dimensions (getter works on all officer versions)
ss <- slide_size(ppt)
slide_w <- ss$width
slide_h <- ss$height

# Choose layouts that exist in the loaded deck
get_layout_pair <- function(ppt, layout_name) {
  lsum <- tryCatch(layout_summary(ppt), error = function(e) NULL)
  if (is.null(lsum)) return(list(layout = layout_name, master = "Office Theme"))
  idx <- which(lsum$layout == layout_name)
  if (length(idx) > 0) return(list(layout = lsum$layout[idx[1]], master = lsum$master[idx[1]]))
  list(layout = lsum$layout[1], master = lsum$master[1])
}

lp_title <- get_layout_pair(ppt, "Title Slide")
lp_tc    <- get_layout_pair(ppt, "Title and Content")

# Slide helper with full-width footnote
add_chart_slide <- function(ppt, title, subtitle, img_path, footer_text) {
  
  ppt <- add_slide(ppt, layout = lp_tc$layout, master = lp_tc$master)
  
  left <- 0.4
  width <- slide_w - 0.8
  
  # Title
  ppt <- ph_with(
    ppt, value = title,
    location = ph_location(left = left, top = 0.15, width = width, height = 0.5)
  )
  
  # Subtitle
  ppt <- ph_with(
    ppt, value = subtitle,
    location = ph_location(left = left, top = 0.65, width = width, height = 0.35)
  )
  
  # Image box
  foot_h   <- 0.80
  foot_top <- slide_h - foot_h - 0.10
  img_top  <- 1.05
  img_h    <- max(1, foot_top - img_top - 0.10)
  
  if (file.exists(img_path)) {
    ppt <- ph_with(
      ppt,
      value = external_img(img_path, width = width, height = img_h),
      location = ph_location(left = left, top = img_top, width = width, height = img_h)
    )
  } else {
    ppt <- ph_with(
      ppt,
      value = paste0("Missing image: ", img_path),
      location = ph_location(left = left, top = img_top, width = width, height = img_h)
    )
  }
  
  # Full-width footnote
  ppt <- ph_with(
    ppt, value = footer_text,
    location = ph_location(left = left, top = foot_top, width = width, height = foot_h)
  )
  
  ppt
}

# Title slide
ppt <- add_slide(ppt, layout = lp_title$layout, master = lp_title$master)
ppt <- ph_with(ppt, "Euro area 10 year sovereign yield spillovers",
               location = ph_location(left = 0.6, top = 1.6, width = slide_w - 1.2, height = 0.8))
ppt <- ph_with(ppt,
               paste0("Rolling Diebold-Yilmaz spillovers, VAR(", p_use, "), daily data, window ", window, ", horizon ", H, "\n",
                      "Countries: US, JP, UK, DE, FR, IT, ES"),
               location = ph_location(left = 0.6, top = 2.6, width = slide_w - 1.2, height = 0.8))

# Executive summary slide
ppt <- add_slide(ppt, layout = lp_tc$layout, master = lp_tc$master)
ppt <- ph_with(ppt, "Executive summary", location = ph_location(left = 0.6, top = 0.3, width = slide_w - 1.2, height = 0.6))
ppt <- ph_with(
  ppt,
  paste0(
    "This deck summarises how shocks transmit across sovereign 10 year yields, with a focus on the euro area.\n\n",
    "Key themes\n",
    ". External versus within-EA drivers are time-varying and regime-dependent.\n",
    ". Fragmentation episodes coincide with stronger within-EA transmission and sharper asymmetries.\n",
    ". US and UK shocks dominate in specific regimes, while Japan becomes relevant around BoJ policy shifts.\n",
    ". Within the EA, the identity of net transmitters rotates across time.\n\n",
    "Event markers use numbered IDs. See events_key.csv in the chart_data folder."
  ),
  location = ph_location(left = 0.6, top = 1.1, width = slide_w - 1.2, height = slide_h - 1.4)
)

# Methodology slide
ppt <- add_slide(ppt, layout = lp_tc$layout, master = lp_tc$master)
ppt <- ph_with(ppt, "Methodology overview (intuitive)", location = ph_location(left = 0.6, top = 0.3, width = slide_w - 1.2, height = 0.6))
ppt <- ph_with(
  ppt,
  paste0(
    "1) Model\n",
    "Rolling VAR(", p_use, ") estimated on daily 10 year yields in levels.\n\n",
    "2) Spillovers\n",
    "Diebold-Yilmaz FEVD spillover table at horizon ", H, ".\n\n",
    "3) Aggregation\n",
    "EA block = DE, FR, IT, ES. External block = US, JP, UK. Average spillover into an EA member is bounded 0-100.\n\n",
    "4) Presentation\n",
    "Daily series retained but smoothed with a ", k_spill_smooth, "-day centred moving average for readability."
  ),
  location = ph_location(left = 0.6, top = 1.1, width = slide_w - 1.2, height = slide_h - 1.4)
)

# Chart slides
ppt <- add_chart_slide(ppt, "Long run yield regimes", "Secular decline through QE, then post-2021 reversal",
                       file.path(charts_dir, "chart01_longrun_yields_sm63.png"),
                       paste0("Footnote, yields smoothed (", k_yields_long, " days). Colours fixed per country. Event IDs match events_key.csv."))

ppt <- add_chart_slide(ppt, "Euro area fragmentation", "Spreads to Germany highlight stress and fragmentation episodes",
                       file.path(charts_dir, "chart02_EA_spreads_to_DE_bp_sm63.png"),
                       paste0("Footnote, spreads in basis points, smoothed (", k_yields_long, " days)."))

ppt <- add_chart_slide(ppt, "Recent yield regime", "Germany and Italy versus US, UK, Japan",
                       file.path(charts_dir, "chart03_recent_yields_DE_IT_outside_sm21.png"),
                       paste0("Footnote, yields smoothed (", k_yields_short, " days)."))

ppt <- add_chart_slide(ppt, "Fragmentation and dispersion", "Standardised measures highlight regime shifts",
                       file.path(charts_dir, "chart04_fragmentation_vs_dispersion_z.png"),
                       "Footnote, standardised indices from smoothed series.")

ppt <- add_chart_slide(ppt, "Spillovers into EA", "Outside versus within EA",
                       file.path(charts_dir, "chart05_spill_into_EA_outside_vs_within.png"),
                       paste0("Footnote, rolling DY spillovers VAR(", p_use, "), window ", window, ", horizon ", H, ", smoothed ", k_spill_smooth, " days."))

ppt <- add_chart_slide(ppt, "External spillovers into EA", "US vs UK vs Japan",
                       file.path(charts_dir, "chart06_external_sources_into_EA.png"),
                       "Footnote, average spillover into an EA member by external origin.")

ppt <- add_chart_slide(ppt, "Within-EA sources into EA", "2x2 panels by origin",
                       file.path(charts_dir, "chart07_internal_sources_into_EA_2x2.png"),
                       "Footnote, internal origin contributions into EA receivers, smoothed.")

ppt <- add_chart_slide(ppt, "Net transmitters within EA", "Directionality inside EA",
                       file.path(charts_dir, "chart08_net_withinEA_2x2.png"),
                       "Footnote, net within-EA transmitter per bilateral link, smoothed.")

ppt <- add_chart_slide(ppt, "Incoming to Germany", "Within-EA bilateral contributions into DE",
                       file.path(charts_dir, "chart_incoming_withinEA_to_DE.png"),
                       "Footnote, bilateral origin-to-receiver spillovers, smoothed.")

ppt <- add_chart_slide(ppt, "Incoming to France", "Within-EA bilateral contributions into FR",
                       file.path(charts_dir, "chart_incoming_withinEA_to_FR.png"),
                       "Footnote, bilateral origin-to-receiver spillovers, smoothed.")

ppt <- add_chart_slide(ppt, "Incoming to Italy", "Within-EA bilateral contributions into IT",
                       file.path(charts_dir, "chart_incoming_withinEA_to_IT.png"),
                       "Footnote, bilateral origin-to-receiver spillovers, smoothed.")

ppt <- add_chart_slide(ppt, "Incoming to Spain", "Within-EA bilateral contributions into ES",
                       file.path(charts_dir, "chart_incoming_withinEA_to_ES.png"),
                       "Footnote, bilateral origin-to-receiver spillovers, smoothed.")

ppt <- add_chart_slide(ppt, "Regime summary", "2015-2019, 2020-2021, 2022-2024, 2024-present",
                       chart13_png,
                       "Footnote, regime means of spillovers into EA by origin, buffered end date to avoid trailing gaps.")

# Closing slide
ppt <- add_slide(ppt, layout = lp_tc$layout, master = lp_tc$master)
ppt <- ph_with(ppt, "Concluding interpretation and writing angles",
               location = ph_location(left = 0.6, top = 0.3, width = slide_w - 1.2, height = 0.6))
ppt <- ph_with(
  ppt,
  paste0(
    "Suggested angles for the note\n",
    ". Regimes, secular QE era versus post-2021 inflation shock.\n",
    ". Fragmentation, spreads and dispersion as a state variable that changes transmission.\n",
    ". External dominance, when US or UK shocks explain more of EA variance.\n",
    ". Internal dominance, rotation in net transmitters inside EA.\n",
    ". Event narrative, align spillover changes with monetary and fiscal shocks using the event markers.\n\n",
    "All chart data are written to the chart_data folder alongside the PNG files."
  ),
  location = ph_location(left = 0.6, top = 1.1, width = slide_w - 1.2, height = slide_h - 1.4)
)

print(ppt, target = ppt_file)

cat("\nOutputs written\n")
cat("Charts folder: ", charts_dir, "\n", sep = "")
cat("Chart data folder: ", data_dir, "\n", sep = "")
cat("PowerPoint file: ", ppt_file, "\n", sep = "")
