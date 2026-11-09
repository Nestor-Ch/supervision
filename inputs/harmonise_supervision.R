# =============================================================================
# harmonise_supervision.R
#
# Recodes the supervision survey (ODK Central export) into the harmonised
# variables defined in supervision_survey_educators_2_harmonised.xlsx.
#
#   1. split the data by population (teachers / heads + deputies)
#   2. recode each population's source columns into the h_* variables,
#      driven by the form's `recode_map` sheet
#   3. stack the two populations into one dataframe
#
# Output format matches the ODK Central export: select_one = one code,
# select_multiple = space-separated codes (ordered as in the choices sheet).
# Optional: 0/1 dummy columns per option (h_var/option), see MAKE_DUMMIES.
#
# Requires: readxl
# =============================================================================

library(readxl)

# ---- paths ------------------------------------------------------------------
data_path <- "supervision_df_02_11.xlsx"
form_path <- "supervision_survey_educators_2_harmonised.xlsx"
out_xlsx   <- "supervision_df_02_11_harmonised.xlsx"

MAKE_DUMMIES <- FALSE   # TRUE adds h_var/option 0-1 columns for select_multiple

# ---- inputs -----------------------------------------------------------------
df  <- read_excel(data_path, col_types = "text")
srv <- read_excel(form_path, sheet = "survey",     col_types = "text")
chc <- read_excel(form_path, sheet = "choices",    col_types = "text")
map <- read_excel(form_path, sheet = "recode_map", col_types = "text")
names(map)[names(map) == "source_code(s)"] <- "source_codes"

h_vars <- srv$name
h_type <- setNames(sub(" .*", "", srv$type), srv$name)   # select_one / select_multiple

# which q_posada values belong to each population
pop_values <- list(teacher = "teacher", manager = c("head", "deputy"))

# ---- helpers ----------------------------------------------------------------
# split "a | b | c" from the recode map into codes
map_codes <- function(s) trimws(strsplit(s, "\\|")[[1]])

# TRUE/FALSE/NA: is `code` among the space-separated codes of x
has_code <- function(x, code) {
  out <- vapply(strsplit(x, "\\s+"), function(cs) code %in% cs, logical(1))
  out[is.na(x)] <- NA
  out
}

# source code -> harmonised code lookup for one variable / population
lookup_for <- function(hv, pop) {
  m <- map[map$h_variable == hv & map$population == pop, ]
  if (nrow(m) == 0) return(NULL)
  lk <- do.call(rbind, lapply(seq_len(nrow(m)), function(i)
    data.frame(src = map_codes(m$source_codes[i]), h = m$h_code[i],
               stringsAsFactors = FALSE)))
  list(source = unique(m$source_variable), lookup = lk)
}

warn_unmapped <- function(hv, pop, codes) {
  codes <- setdiff(unique(codes), NA)
  if (length(codes))
    warning(sprintf("%s (%s): source codes not in recode_map, set to NA: %s",
                    hv, pop, paste(codes, collapse = ", ")), call. = FALSE)
}

# generic recode, driven entirely by recode_map
recode_generic <- function(d, hv, pop) {
  spec <- lookup_for(hv, pop)
  if (is.null(spec)) return(rep(NA_character_, nrow(d)))   # not asked of this population
  stopifnot(length(spec$source) == 1)
  x  <- d[[spec$source]]
  lk <- spec$lookup

  if (h_type[[hv]] == "select_one") {
    out <- lk$h[match(x, lk$src)]
    warn_unmapped(hv, pop, x[!is.na(x) & is.na(out)])
    return(out)
  }

  # select_multiple: map each code, de-duplicate (composites), keep choice order
  h_order <- chc$name[chc$list_name == hv]
  parts   <- strsplit(x, "\\s+")
  warn_unmapped(hv, pop, setdiff(unlist(parts), lk$src))
  vapply(parts, function(cs) {
    if (all(is.na(cs))) return(NA_character_)
    h <- unique(lk$h[match(cs, lk$src)])
    h <- h_order[h_order %in% h]
    if (length(h) == 0) NA_character_ else paste(h, collapse = " ")
  }, character(1))
}

# variables that need logic beyond a one-to-one code lookup
recode_one <- function(d, hv, pop) {
  if (hv == "h_population") return(rep(pop, nrow(d)))

  if (hv == "h_sought" && pop == "teacher") {
    # yes = took part, or ordered / tried to order; no = neither
    yes <- d$q_participated %in% "yes" |
           d$q_tried_order  %in% c("yes_success", "yes_failed")
    no  <- d$q_participated %in% "no" & d$q_tried_order %in% "no"
    return(ifelse(yes, "yes", ifelse(no, "no", NA_character_)))
  }

  if (hv == "h_admin_informs") {
    # teachers: 'administration' among info sources
    # managers: 'informs' among admin support (NA where not asked)
    flag <- if (pop == "teacher") has_code(d$q_info_sources, "administration")
            else                  has_code(d$q_admin_support, "informs")
    return(ifelse(flag, "yes", "no"))
  }

  recode_generic(d, hv, pop)
}

# ---- 1. split, 2. recode ----------------------------------------------------
harmonise_population <- function(pop) {
  d   <- df[df$q_posada %in% pop_values[[pop]], ]
  out <- data.frame(KEY = d$KEY, stringsAsFactors = FALSE)
  for (hv in h_vars) out[[hv]] <- recode_one(d, hv, pop)
  out
}

teachers <- harmonise_population("teacher")
managers <- harmonise_population("manager")

# ---- 3. stack ---------------------------------------------------------------
harm <- rbind(teachers, managers)

dropped <- sum(!df$q_posada %in% unlist(pop_values))
if (dropped) warning(dropped, " rows with unexpected q_posada were dropped.", call. = FALSE)

# ---- optional dummy columns -------------------------------------------------
if (MAKE_DUMMIES) {
  for (hv in h_vars[h_type[h_vars] == "select_multiple"]) {
    for (opt in chc$name[chc$list_name == hv]) {
      harm[[paste0(hv, "/", opt)]] <- as.integer(has_code(harm[[hv]], opt))
    }
  }
}

# ---- checks -----------------------------------------------------------------
cat("Rows by population:\n"); print(table(harm$h_population, useNA = "ifany"))
cat("\nAnswered (non-NA) by population:\n")
print(sapply(split(harm[h_vars], harm$h_population),
             function(x) colSums(!is.na(x))))

# ---- save -------------------------------------------------------------------
writexl::write_xlsx(harm, out_xlsx)
