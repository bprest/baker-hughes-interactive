# =============================================================================
# Baker Hughes North America Rig Count — Interactive Visualization
# =============================================================================
# Sources
#   PRIMARY   : auto-scraped live Excel from bakerhughesrigcount.gcs-web.com
#   HISTORICAL: local Excel file covering 2013–Aug 2025 (user-supplied)
#
# Chart views (toggle via buttons):
#   1. US vs Canada            — total weekly rigs (lines)
#   2. US by rig type          — Oil / Gas / Miscellaneous (lines)
#   3. US by location          — Land / Offshore / Inland Waters (lines)
#   4. US by basin             — all 14 named basins + Other (stacked area)
#   5. US basin — Oil-directed — basins filtered to Oil DrillFor
#   6. US basin — Gas-directed — basins filtered to Gas DrillFor
#   7. US by state             — top N states + Other (stacked area)
#   8. Canada by province      — top N provinces + Other (stacked area)
#
# Time span: range-selector buttons (1Y/2Y/5Y/8Y/10Y/All) + drag rangeslider
# Y-axis: double-click chart to reset; drag y-axis edge to rescale
# =============================================================================

# ── 0. USER CONFIG ────────────────────────────────────────────────────────────

# Path to your historical Excel file (2013–Aug 2025).
# Can be absolute or relative to your R working directory.
HIST_FILE <- "08-29-2025 North America Rig Count Report.xlsx"

# How many US states / CA provinces to show individually (rest → "Other")
TOP_N_US <- 8
TOP_N_CA <- 5

# Output file (written to working directory)
OUT_HTML <- "bh_rig_count_interactive.html"

# Known-good fallback URL (used if auto-scrape fails)
FALLBACK_URL <- "https://bakerhughesrigcount.gcs-web.com/static-files/99401d53-5549-42c3-96fa-844237f73a74"

# ── 1. Dependencies ───────────────────────────────────────────────────────────
required <- c("httr", "rvest", "jsonlite", "readxl",
              "dplyr", "tidyr", "lubridate", "plotly", "stringr")
new_pkgs <- setdiff(required, rownames(installed.packages()))
if (length(new_pkgs)) {
  message("Installing: ", paste(new_pkgs, collapse = ", "))
  install.packages(new_pkgs, repos = "https://cloud.r-project.org")
}
suppressPackageStartupMessages({
  library(httr);     library(rvest);    library(jsonlite)
  library(readxl);   library(dplyr);    library(tidyr)
  library(lubridate); library(plotly);  library(stringr)
})

# ── 2. Auto-discover and download the latest Excel file ───────────────────────
# The BH rig count page is a React SPA. File links appear as UUIDs inside a
# __NEXT_DATA__ JSON blob or raw href attributes. We collect ALL candidate
# UUIDs from the page, then probe each one until we download a valid xlsx.
# If the page is unreachable we fall back to the hardcoded UUID.

PAGE_URL     <- "https://bakerhughesrigcount.gcs-web.com/na-rig-count/"
BASE_URL     <- "https://bakerhughesrigcount.gcs-web.com"
UUID_PATTERN <- "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}"

BROWSER_UA <- paste0("Mozilla/5.0 (Windows NT 10.0; Win64; x64) ",
                     "AppleWebKit/537.36 (KHTML, like Gecko) ",
                     "Chrome/124.0.0.0 Safari/537.36")

# TRUE if `path` starts with PK magic bytes (xlsx = zip archive)
is_valid_xlsx <- function(path) {
  if (!file.exists(path) || file.size(path) < 1000) return(FALSE)
  tryCatch({
    con   <- file(path, "rb")
    magic <- readBin(con, "raw", n = 4)
    close(con)
    identical(magic, as.raw(c(0x50, 0x4b, 0x03, 0x04)))
  }, error = function(e) FALSE)
}

# Extract the best Excel URL from the page HTML.
# Priority order:
#   1. <a> tags whose visible text contains "North America" and "Rig Count"
#      (i.e. the "North America Rig Count Report" link, NOT "Read More")
#   2. Any static-files UUID that appears in the page
#   3. Hardcoded fallback UUID
find_excel_url <- function(html) {
  page <- tryCatch(rvest::read_html(html), error = function(e) NULL)
  
  if (!is.null(page)) {
    # All <a> elements that link to a static-files UUID
    anchors  <- rvest::html_elements(page, paste0("a[href*='static-files']"))
    hrefs    <- rvest::html_attr(anchors, "href")
    texts    <- rvest::html_text2(anchors)
    
    if (length(hrefs) > 0) {
      # Score each link: prefer ones whose text matches the report label
      label_match <- grepl("north.?america", texts, ignore.case = TRUE) &
        grepl("rig.?count",     texts, ignore.case = TRUE)
      
      # Return the first label-matched link, else first static-files link
      matched <- hrefs[label_match]
      chosen  <- if (length(matched) > 0) matched[1] else hrefs[1]
      if (!grepl("^http", chosen)) chosen <- paste0(BASE_URL, chosen)
      
      label <- if (length(matched) > 0) texts[label_match][1] else texts[1]
      message("  Selected link: \"", trimws(label), "\"")
      return(chosen)
    }
  }
  
  # Fallback: raw regex for any static-files UUID in the page
  sf_uuids <- regmatches(html,
                         gregexpr(paste0("(?<=static-files/)", UUID_PATTERN), html, perl = TRUE))[[1]]
  if (length(sf_uuids) > 0) {
    message("  Selected via raw UUID scan (no matching anchor found)")
    return(paste0(BASE_URL, "/static-files/", sf_uuids[1]))
  }
  
  NULL   # signal failure; caller will use FALLBACK_URL
}

LIVE_FILE <- file.path(tempdir(), "bh_rig_count_live.xlsx")
LIVE_URL  <- NULL

message("-- Step 1: Fetching BH page to discover Excel URL...")
page_resp <- tryCatch(
  httr::GET(PAGE_URL,
            httr::user_agent(BROWSER_UA),
            httr::add_headers(
              "Accept"          = "text/html,application/xhtml+xml,*/*;q=0.8",
              "Accept-Language" = "en-US,en;q=0.9",
              "Referer"         = "https://www.google.com/"),
            httr::timeout(30)),
  error = function(e) { message("  Page fetch failed: ", e$message); NULL }
)

discovered_url <- NULL
if (!is.null(page_resp) && httr::status_code(page_resp) == 200) {
  html           <- httr::content(page_resp, as = "text", encoding = "UTF-8")
  discovered_url <- find_excel_url(html)
  if (is.null(discovered_url))
    message("  No suitable link found in page; will use fallback URL.")
} else {
  message("  Page status: ",
          if (!is.null(page_resp)) httr::status_code(page_resp) else "unreachable")
}

# Probe list: discovered URL first, hardcoded fallback second
candidate_urls <- unique(c(discovered_url, FALLBACK_URL))

message("-- Step 2: Probing ", length(candidate_urls), " candidate URL(s)...")
for (url in candidate_urls) {
  message("  Trying: ", url)
  dl <- tryCatch(
    httr::GET(url,
              httr::user_agent(BROWSER_UA),
              httr::write_disk(LIVE_FILE, overwrite = TRUE),
              httr::timeout(120)),
    error = function(e) { message("    Download error: ", e$message); NULL }
  )
  code <- if (!is.null(dl)) httr::status_code(dl) else "error"
  sz   <- if (file.exists(LIVE_FILE)) file.size(LIVE_FILE) else 0
  
  if (!is.null(dl) && code == 200 && is_valid_xlsx(LIVE_FILE)) {
    LIVE_URL <- url
    message("  OK: valid xlsx (", round(sz / 1e6, 1), " MB)")
    break
  }
  message("    Not a valid xlsx (HTTP ", code, ", ", sz, " bytes)")
}

if (is.null(LIVE_URL)) {
  stop(
    "\nCould not download a valid Baker Hughes Excel file.\n",
    "Tried ", length(candidate_urls), " URL(s).\n\n",
    "Manual fix: download the file from\n",
    "  https://bakerhughesrigcount.gcs-web.com/na-rig-count/\n",
    "save it anywhere, then at the top of this script set:\n",
    "  LIVE_FILE <- 'C:/path/to/downloaded_file.xlsx'\n",
    "  LIVE_URL  <- 'manual'\n",
    "and comment out the entire Section 2 block."
  )
}
message("Live URL: ", LIVE_URL)

# ── 3. Reader: NAM Weekly sheet → tidy data frame ────────────────────────────
# Schema (verified May 2026): headers on row 11 (skip = 10)
#   A=Country  B=County  C=Basin  D=GOM  E=DrillFor  F=Location
#   G=State/Province  H=Trajectory  I=Year  J=Month
#   K=US_PublishDate  L=Rig Count Value
read_nam_weekly <- function(path) {
  raw <- read_excel(path, sheet = "NAM Weekly", skip = 10,
                    col_names = TRUE, .name_repair = "universal")
  names(raw) <- tolower(gsub("[^A-Za-z0-9]+", "_", names(raw)))
  names(raw) <- gsub("_+$", "", names(raw))
  
  raw %>%
    filter(!is.na(us_publishdate)) %>%
    transmute(
      week       = as.Date(us_publishdate),
      country    = str_to_title(trimws(country)),
      state_prov = str_to_title(trimws(state_province)),
      basin      = trimws(basin),
      drill_for  = trimws(drillfor),
      location   = trimws(location),
      count      = suppressWarnings(as.numeric(rig_count_value))
    ) %>%
    filter(!is.na(count), count >= 0)
}

# ── 4. Load & merge ───────────────────────────────────────────────────────────
message("Reading live file...")
live_data <- read_nam_weekly(LIVE_FILE)
message("  Rows: ", nrow(live_data),
        "  |  Weeks: ", n_distinct(live_data$week),
        "  |  Range: ", min(live_data$week), " -> ", max(live_data$week))

if (!file.exists(HIST_FILE))
  stop("Historical file not found: ", HIST_FILE,
       "\nUpdate HIST_FILE at the top of this script.")

message("Reading historical file: ", HIST_FILE)
hist_data <- read_nam_weekly(HIST_FILE)
message("  Rows: ", nrow(hist_data),
        "  |  Weeks: ", n_distinct(hist_data$week),
        "  |  Range: ", min(hist_data$week), " -> ", max(hist_data$week))

# Live file wins for any overlapping weeks
weekly <- bind_rows(
  hist_data %>% filter(!week %in% unique(live_data$week)),
  live_data
) %>% arrange(week)

message("\nMerged dataset:")
message("  Rows  : ", nrow(weekly))
message("  Weeks : ", n_distinct(weekly$week))
message("  Range : ", min(weekly$week), " -> ", max(weekly$week))

# ── 5. Aggregated views ───────────────────────────────────────────────────────
latest_wk <- max(weekly$week)

# Helper: order factor levels by count in the latest week, highest first
#         so stacked areas go highest-count (bottom) to lowest-count (top)
order_by_latest <- function(df, grp_col) {
  lvls <- df %>%
    filter(week == latest_wk) %>%
    arrange(desc(count)) %>%
    pull(!!sym(grp_col)) %>%
    as.character()
  # Ensure "Other" always goes on top
  lvls <- c(setdiff(lvls, "Other"), "Other")
  factor(df[[grp_col]], levels = lvls)
}

# 5a. Country totals
na_weekly <- bind_rows(
  weekly %>% group_by(week, country) %>% summarise(count = sum(count), .groups = "drop"),
  weekly %>% group_by(week) %>% summarise(count = sum(count), .groups = "drop") %>%
    mutate(country = "North America")
) %>% arrange(week, country)

# 5b. US by rig type
us_by_type <- weekly %>%
  filter(country == "United States") %>%
  group_by(week, drill_for) %>%
  summarise(count = sum(count), .groups = "drop") %>%
  arrange(week, drill_for)

# 5c. US by location
us_by_location <- weekly %>%
  filter(country == "United States") %>%
  group_by(week, location) %>%
  summarise(count = sum(count), .groups = "drop") %>%
  arrange(week, location)

# 5d–5f. Basin helpers
NAMED_BASINS <- c(
  "Permian", "Eagle Ford", "Haynesville", "Marcellus",
  "DJ-Niobrara", "Williston", "Utica", "Ardmore Woodford",
  "Cana Woodford", "Arkoma Woodford", "Barnett",
  "Mississippian", "Fayetteville", "Granite Wash"
)

make_basin_df <- function(data) {
  all_weeks <- sort(unique(weekly$week))   # full merged date spine
  agg <- data %>%
    mutate(basin_grp = if_else(basin %in% NAMED_BASINS, basin, "Other")) %>%
    group_by(week, basin_grp) %>%
    summarise(count = sum(count), .groups = "drop") %>%
    # Fill every (week x basin) combination with 0 where missing.
    # Plotly stacked areas require all traces in a stackgroup to share
    # identical x-values; gaps cause misaligned stacks and broken hover.
    tidyr::complete(week = all_weeks, basin_grp, fill = list(count = 0))
  latest_order <- agg %>%
    filter(week == latest_wk) %>%
    arrange(desc(count)) %>%
    pull(basin_grp) %>%
    as.character()
  latest_order <- c(setdiff(latest_order, "Other"), "Other")
  agg %>%
    mutate(basin_grp = factor(basin_grp, levels = latest_order)) %>%
    arrange(week, basin_grp)
}

# 5d. All US basins
us_by_basin <- weekly %>%
  filter(country == "United States") %>%
  make_basin_df()

# 5e. Oil-directed basins only
us_basin_oil <- weekly %>%
  filter(country == "United States", drill_for == "Oil") %>%
  make_basin_df()

# 5f. Gas-directed basins only
us_basin_gas <- weekly %>%
  filter(country == "United States", drill_for == "Gas") %>%
  make_basin_df()

# 5g. US by state — top N + Other, ordered by latest week
top_us_states <- weekly %>%
  filter(country == "United States", week == latest_wk) %>%
  group_by(state_prov) %>%
  summarise(last_count = sum(count), .groups = "drop") %>%
  slice_max(last_count, n = TOP_N_US) %>%
  pull(state_prov)

us_by_state <- weekly %>%
  filter(country == "United States") %>%
  mutate(state_grp = if_else(state_prov %in% top_us_states, state_prov, "Other")) %>%
  group_by(week, state_grp) %>%
  summarise(count = sum(count), .groups = "drop") %>%
  tidyr::complete(week = sort(unique(weekly$week)), state_grp,
                  fill = list(count = 0)) %>%
  mutate(state_grp = {
    latest_order <- filter(., week == latest_wk) %>%
      arrange(desc(count)) %>% pull(state_grp) %>% as.character()
    latest_order <- c(setdiff(latest_order, "Other"), "Other")
    factor(state_grp, levels = latest_order)
  }) %>%
  arrange(week, state_grp)

# 5h. Canada by province — top N + Other, ordered by latest week
top_ca_provs <- weekly %>%
  filter(country == "Canada", week == latest_wk) %>%
  group_by(state_prov) %>%
  summarise(last_count = sum(count), .groups = "drop") %>%
  slice_max(last_count, n = TOP_N_CA) %>%
  pull(state_prov)

ca_by_prov <- weekly %>%
  filter(country == "Canada") %>%
  mutate(prov_grp = if_else(state_prov %in% top_ca_provs, state_prov, "Other")) %>%
  group_by(week, prov_grp) %>%
  summarise(count = sum(count), .groups = "drop") %>%
  tidyr::complete(week = sort(unique(weekly$week)), prov_grp,
                  fill = list(count = 0)) %>%
  mutate(prov_grp = {
    latest_order <- filter(., week == latest_wk) %>%
      arrange(desc(count)) %>% pull(prov_grp) %>% as.character()
    latest_order <- c(setdiff(latest_order, "Other"), "Other")
    factor(prov_grp, levels = latest_order)
  }) %>%
  arrange(week, prov_grp)

message("Top US states    : ", paste(top_us_states, collapse = ", "))
message("Top CA provinces : ", paste(top_ca_provs,  collapse = ", "))

# ── 6. Colour palette ─────────────────────────────────────────────────────────
PALETTE <- c(
  "United States"    = "#1f77b4",
  "Canada"           = "#2ca02c",
  "North America"    = "#636efa",
  "Oil"              = "#d62728",
  "Gas"              = "#9467bd",
  "Miscellaneous"    = "#8c564b",
  "Land"             = "#bcbd22",
  "Offshore"         = "#17becf",
  "Inland Waters"    = "#aec7e8",
  "Permian"          = "#e6194b",
  "Eagle Ford"       = "#f58231",
  "Haynesville"      = "#3cb44b",
  "Marcellus"        = "#4363d8",
  "DJ-Niobrara"      = "#42d4f4",
  "Williston"        = "#911eb4",
  "Utica"            = "#f032e6",
  "Ardmore Woodford" = "#fabebe",
  "Cana Woodford"    = "#469990",
  "Arkoma Woodford"  = "#e6beff",
  "Barnett"          = "#9a6324",
  "Mississippian"    = "#fffac8",
  "Fayetteville"     = "#800000",
  "Granite Wash"     = "#aaffc3",
  "Texas"            = "#e6194b",
  "Oklahoma"         = "#f58231",
  "Louisiana"        = "#b5bd00",
  "Wyoming"          = "#3cb44b",
  "Colorado"         = "#42d4f4",
  "North Dakota"     = "#4363d8",
  "New Mexico"       = "#911eb4",
  "Pennsylvania"     = "#f032e6",
  "West Virginia"    = "#fabebe",
  "Utah"             = "#469990",
  "Alberta"          = "#e6194b",
  "British Columbia" = "#f58231",
  "Saskatchewan"     = "#ffe119",
  "Manitoba"         = "#3cb44b",
  "Nl Offshore"      = "#42d4f4",
  "Other"            = "#aaaaaa"
)

FALLBACK <- c("#e377c2","#7f7f7f","#bcbd22","#17becf",
              "#8c564b","#aec7e8","#c49c94","#dbdb8d","#9edae5","#ad494a")

get_color <- function(nms) {
  sapply(seq_along(nms), function(i) {
    v <- PALETTE[nms[i]]
    if (!is.na(v)) v else FALLBACK[((i - 1) %% length(FALLBACK)) + 1]
  })
}

# ── 7. Trace builder ──────────────────────────────────────────────────────────
make_traces <- function(df, x_col, grp_col, y_col,
                        visible = TRUE, mode = "lines", stackgroup = NULL) {
  grps <- levels(factor(df[[grp_col]]))
  cols <- get_color(grps)
  lapply(seq_along(grps), function(i) {
    g <- grps[i]; col <- cols[i]
    d <- df[df[[grp_col]] == g, ]
    tr <- list(
      x = d[[x_col]], y = d[[y_col]],
      type = "scatter", mode = mode, name = g,
      line   = list(color = col, width = 1.6),
      marker = list(color = col, size = 4),
      visible  = visible,
      legendgroup = g,
      # Hidden traces must be skipped by hover entirely; shown traces get
      # full tooltip. vis_for() updates this alongside visible= on each switch.
      hoverinfo     = if (isTRUE(visible)) "x+y+name" else "skip",
      hovertemplate = paste0(
        "<b>", g, "</b><br>%{x|%b %d, %Y}<br>Rigs: %{y:.0f}<extra></extra>")
    )
    if (!is.null(stackgroup)) {
      tr$stackgroup <- stackgroup
      tr$fillcolor  <- paste0(col, "99")
      tr$line$width <- 0.4
    }
    tr
  })
}

# ── 8. Build all trace groups ─────────────────────────────────────────────────
tr_A <- make_traces(filter(na_weekly, country %in% c("United States", "Canada")),
                    "week", "country",   "count", visible = TRUE)
tr_B <- make_traces(us_by_type,     "week", "drill_for", "count", visible = FALSE)
tr_C <- make_traces(us_by_location, "week", "location",  "count", visible = FALSE)
tr_D <- make_traces(us_by_basin,    "week", "basin_grp", "count",
                    visible = FALSE, stackgroup = "basins_all")
tr_E <- make_traces(us_basin_oil,   "week", "basin_grp", "count",
                    visible = FALSE, stackgroup = "basins_oil")
tr_F <- make_traces(us_basin_gas,   "week", "basin_grp", "count",
                    visible = FALSE, stackgroup = "basins_gas")
tr_G <- make_traces(us_by_state,    "week", "state_grp", "count",
                    visible = FALSE, stackgroup = "us_states")
tr_H <- make_traces(ca_by_prov,     "week", "prov_grp",  "count",
                    visible = FALSE, stackgroup = "ca_provs")

all_traces <- c(tr_A, tr_B, tr_C, tr_D, tr_E, tr_F, tr_G, tr_H)
ns <- c(length(tr_A), length(tr_B), length(tr_C), length(tr_D),
        length(tr_E), length(tr_F), length(tr_G), length(tr_H))
N  <- sum(ns)

vis_for <- function(idx) {
  ends   <- cumsum(ns)
  starts <- c(1, ends[-length(ends)] + 1)
  v <- rep(FALSE, N)
  for (i in idx) v[starts[i]:ends[i]] <- TRUE
  v
}

# Hidden stacked traces keep their stackgroup active in Plotly's hover engine
# even when visible=FALSE, causing tooltips to misfire on oil/gas basin views.
# Setting hoverinfo="skip" on hidden traces removes them from hover entirely.
hover_for <- function(idx) {
  ifelse(vis_for(idx), "x+y+name", "skip")
}

# ── 9. Assemble figure ────────────────────────────────────────────────────────
make_btn <- function(label, vis, hover, subtitle) {
  list(label = label, method = "update",
       args = list(
         list(visible   = vis,
              hoverinfo = hover),
         list(
           yaxis = list(title = "Active Rigs", autorange = TRUE)
         )
       ))
}

fig <- plot_ly()
for (tr in all_traces) fig <- do.call(add_trace, c(list(p = fig), tr))

fig <- fig %>%
  layout(
    # Title is rendered as a plain HTML div ABOVE the widget via
    # htmltools::prependContent — completely outside Plotly's coordinate
    # system, so it never overlaps buttons regardless of browser width.
    margin = list(t = 130, b = 80, l = 60, r = 20),
    
    xaxis = list(
      title = "",
      # Range selector removed — rangeslider + 1Y/2Y/5Y/10Y/All buttons
      # below handle all time-span selection needs without eating header space
      rangeslider = list(visible = TRUE, thickness = 0.05)
    ),
    
    yaxis = list(
      title      = "Active Rigs",
      fixedrange = FALSE,
      autorange  = TRUE
    ),
    
    # hovermode "closest" is per-trace and works correctly with stacked areas
    # that have hidden sibling traces. "x unified" breaks when hidden stacked
    # traces from other views share the same x domain.
    hovermode = "closest",
    legend    = list(orientation = "h", y = -0.18, yanchor = "top",
                     font = list(size = 11)),
    paper_bgcolor = "#f9f9f9",
    plot_bgcolor  = "#ffffff",
    
    # ── View-toggle buttons (top row) ───────────────────────────────────────
    updatemenus = list(list(
      type        = "buttons",
      direction   = "right",
      x           = 0,
      y           = 1.13,
      xanchor     = "left",
      yanchor     = "top",
      pad         = list(r = 5, t = 5),
      showactive  = TRUE,
      bgcolor     = "#e8e8e8",
      bordercolor = "#bbbbbb",
      font        = list(size = 11),
      buttons = list(
        make_btn("US vs Canada",
                 vis_for(1), hover_for(1), "North America — weekly"),
        make_btn("US by Type",
                 vis_for(2), hover_for(2), "US rigs by drill target — weekly"),
        make_btn("US by Location",
                 vis_for(3), hover_for(3), "US rigs by location — weekly"),
        make_btn("US by Basin (All)",
                 vis_for(4), hover_for(4), "US rigs by basin, all targets — stacked"),
        make_btn("US Basin — Oil",
                 vis_for(5), hover_for(5), "US rigs by basin, oil-directed — stacked"),
        make_btn("US Basin — Gas",
                 vis_for(6), hover_for(6), "US rigs by basin, gas-directed — stacked"),
        make_btn(paste0("US by State (top ", TOP_N_US, ")"),
                 vis_for(7), hover_for(7),
                 paste0("US top ", TOP_N_US, " states + Other — stacked")),
        make_btn(paste0("Canada by Prov. (top ", TOP_N_CA, ")"),
                 vis_for(8), hover_for(8),
                 paste0("Canada top ", TOP_N_CA, " provinces + Other — stacked"))
      )
    ),
    # ── Time-span buttons (below view buttons, above plot) ──────────────────
    list(
      type      = "buttons",
      direction = "right",
      x = 0, y = 1.0, xanchor = "left", yanchor = "top",
      pad = list(r = 4, t = 4),
      showactive = TRUE,
      bgcolor = "#e8e8e8", bordercolor = "#bbbbbb",
      font = list(size = 10),
      buttons = list(
        list(label = "1Y",  method = "relayout",
             args = list(list(xaxis.range = list(
               as.character(max(weekly$week) - 365),
               as.character(max(weekly$week)))))),
        list(label = "2Y",  method = "relayout",
             args = list(list(xaxis.range = list(
               as.character(max(weekly$week) - 730),
               as.character(max(weekly$week)))))),
        list(label = "5Y",  method = "relayout",
             args = list(list(xaxis.range = list(
               as.character(max(weekly$week) - 1826),
               as.character(max(weekly$week)))))),
        list(label = "8Y",  method = "relayout",
             args = list(list(xaxis.range = list(
               as.character(max(weekly$week) - 2922),
               as.character(max(weekly$week)))))),
        list(label = "10Y", method = "relayout",
             args = list(list(xaxis.range = list(
               as.character(max(weekly$week) - 3652),
               as.character(max(weekly$week)))))),
        list(label = "All", method = "relayout",
             args = list(list(xaxis.autorange = TRUE)))
      )
    )),
    
    annotations = list(
      # Title — sits in margin above the view buttons
      list(
        text      = paste0("<b>Baker Hughes North American Rig Count</b> - weekly, 2013 to ", gsub('  ', ' ',(format(as.Date(max(weekly$week)),  "%B %e, %Y")))),
        font      = list(size = 14, color = "#333333"),
        showarrow = FALSE, xref = "paper", yref = "paper",
        x = 0, y = 1.155, xanchor = "left", yanchor = "bottom",
        align = "left"
      ),
      # Footer source note
      list(
        text = paste0(
          "Source: Baker Hughes  |  ",
          "Historical file: Jan 2013–Aug 2025  |  ",
          "Live download through: ", max(weekly$week),
          "  |  Stacked basins/states sorted highest→lowest as of latest week"
        ),
        showarrow = FALSE, xref = "paper", yref = "paper",
        x = 0, y = -0.15, xanchor = "left", yanchor = "top",
        font = list(size = 9, color = "#999999")
      )
    )
  ) %>%
  config(
    # scrollZoom lets the user zoom y-axis with scroll wheel when hovering over it
    scrollZoom  = TRUE,
    displayModeBar = TRUE,
    modeBarButtonsToRemove = c("select2d", "lasso2d"),
    toImageButtonOptions = list(
      format = "png", filename = "bh_rig_count",
      width = 1400, height = 750, scale = 2)
  )

# ── 10. Save & open ───────────────────────────────────────────────────────────
# Single self-contained HTML — all JS/CSS inlined via pandoc.
# Title is a Plotly annotation sitting in the margin above the buttons.
htmlwidgets::saveWidget(fig, OUT_HTML, selfcontained = TRUE)
message("\nChart saved -> ", normalizePath(OUT_HTML))
message("File size: ", round(file.size(OUT_HTML) / 1e6, 1), " MB")
browseURL(OUT_HTML)

# ── 11. Console summary ───────────────────────────────────────────────────────
us_n <- weekly %>% filter(week == latest_wk, country == "United States") %>%
  summarise(n = sum(count)) %>% pull(n)
ca_n <- weekly %>% filter(week == latest_wk, country == "Canada") %>%
  summarise(n = sum(count)) %>% pull(n)

message("\n=== Latest week (", latest_wk, ") ===")
message("  US     : ", us_n, " rigs")
message("  Canada : ", ca_n, " rigs")
message("  NAM    : ", us_n + ca_n, " rigs")
message("\n=== Full merged history ===")
message("  Weeks  : ", n_distinct(weekly$week))
message("  Range  : ", min(weekly$week), " -> ", max(weekly$week))
message("\nLive Excel URL used: ", LIVE_URL)