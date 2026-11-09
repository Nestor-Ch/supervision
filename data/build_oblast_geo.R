# Rebuilds data/oblast_geo.geojson (cached oblast polygons; see CLAUDE.md "Maps").
# Run locally only - the app never calls rnaturalearth at runtime.
library(rnaturalearth); library(sf); library(dplyr)

ua <- ne_states(country = "Ukraine", returnclass = "sf")
# Natural Earth files Crimea and Sevastopol under Russia (de facto control);
# the map shows Ukraine's internationally recognized territory.
crimea <- ne_states(country = "Russia", returnclass = "sf") %>%
  filter(name %in% c("Crimea", "Sevastopol"))

geo <- bind_rows(ua, crimea) %>%
  mutate(
    # Natural Earth's own name_uk labels the oblast "Київ"; survey calls it "Київська"
    name = case_when(
      name == "Kiev"      ~ "Київська",
      name == "Kiev City" ~ "Київ",
      TRUE                ~ sub(" область$", "", name_uk)
    )
  ) %>%
  select(name)

st_write(geo, "data/oblast_geo.geojson", delete_dsn = TRUE, quiet = TRUE)
print(st_drop_geometry(geo))
