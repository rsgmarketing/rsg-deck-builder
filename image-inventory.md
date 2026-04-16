# Image Inventory

All images are in `assets/` relative to this folder. Use these exact paths in python-pptx scripts.

## Selection Rules

1. **Match brand to content.** Norlake deck = Norlake products and logos. Master-Bilt deck = Master-Bilt products and logos. Cross-brand or RSG corporate = RSG logos.
2. **One hero image per slide maximum.** Don't crowd slides with multiple product shots.
3. **Use transparent-background PNGs for overlays.** The capsule-pak-eco-transparent.png is ideal for compositing over colored backgrounds.
4. **Dark background slides get white logos.** Light background slides get dark logos. Never reverse this.
5. **No oval logos.** Always use the "no oval" Norlake variants for presentations.
6. **Application shots over product shots when telling a story.** Use the kitchen-setting or c-store images when the slide is about the customer's world, not the product spec.

---

## Logos

All in `assets/logos/`.

### Norlake
| File | Background | Notes |
|---|---|---|
| `norlake-white.png` | Dark backgrounds | Full Norlake logo, white |
| `norlake-dark.png` | Light backgrounds | Full Norlake logo, PMS 295 navy |
| `norlake-no-oval-white.png` | Dark backgrounds | Preferred for presentations — cleaner |
| `norlake-no-oval-dark.png` | Light backgrounds | Preferred for presentations — cleaner |

### Master-Bilt
| File | Background | Notes |
|---|---|---|
| `masterbilt-white.png` | Dark backgrounds | MB logo, white |
| `masterbilt-dark.png` | Light backgrounds | MB logo, full color |

### RSG Corporate
| File | Background | Notes |
|---|---|---|
| `rsg-white.png` | Dark backgrounds | RSG only, no brand names |
| `rsg-dark.png` | Light backgrounds | RSG only, no brand names |
| `rsg-3brands-white.png` | Dark backgrounds | RSG + Norlake + Master-Bilt lockup |
| `rsg-3brands-dark.png` | Light backgrounds | RSG + Norlake + Master-Bilt lockup |
| `rsg-3brands-white-rsg.png` | Dark with blue accent | White RSG text variant |
| `rsg-nl-mb-horizontal-dark.png` | Light backgrounds | Horizontal RSG + NL + MB layout |

### Norlake Scientific
| File | Background | Notes |
|---|---|---|
| `norlake-scientific-white.png` | Dark backgrounds | No oval variant, white |
| `norlake-scientific-dark.png` | Light backgrounds | No oval variant, black |

### Certification
| File | Use For |
|---|---|
| `iso-9001.png` | Quality certification slides |
| `iso-14001.png` | Environmental certification slides |

---

## Product Images

### Norlake — Walk-Ins (`assets/products/norlake/walk-ins/`)

| File | Description | Best For |
|---|---|---|
| `kold-locker-capsule-pak-left.jpg` | Kold Locker with Capsule Pak ECO on top, left angle | Hero shot for Kold Locker, self-contained walk-in |
| `kold-locker-capsule-pak-right.jpg` | Same unit, right angle | Alternate angle when left doesn't fit layout |
| `fast-trak.jpg` | Standard smooth-panel Fast-Trak walk-in | Fast-Trak promos, stock walk-in, quick-ship messaging |
| `fast-trak-combo.jpg` | Fast-Trak combination cooler/freezer | Multi-unit configurations, combo messaging |
| `fineline.jpg` | FineLine custom walk-in | Custom/premium walk-in content |
| `foodservice.jpg` | Generic foodservice walk-in | General walk-in content, non-product-specific slides |
| `indoor-stainless.jpg` | Stainless steel interior walk-in | Premium/stainless promos, hygiene messaging |
| `interior-shelving.jpg` | Walk-in interior with shelving and product | Interior capacity, storage capability |
| `kitchen-setting-left.jpg` | Walk-in in kitchen setting (render) | Lifestyle/application context, customer-facing decks |

### Norlake — Refrigeration Systems (`assets/products/norlake/refrigeration/`)

| File | Description | Best For |
|---|---|---|
| `capsule-pak-eco-left.jpg` | Capsule Pak ECO unit, left angle | Self-contained refrigeration hero shot |
| `capsule-pak-eco-transparent.png` | Capsule Pak ECO, transparent background | Overlays on colored backgrounds, composite layouts |
| `capsule-pak-eco-outdoor.jpg` | Capsule Pak in outdoor install | Outdoor applications, weatherproof messaging |
| `controller.jpg` | LogiTemp controller closeup | Controller/monitoring content |

### Norlake — Electronic Controllers (`assets/products/norlake/controllers/`)

| File | Description | Best For |
|---|---|---|
| `logitemp-laptop.jpg` | LogiTemp with monitoring laptop | Remote monitoring, IoT messaging |
| `logitemp-touch.jpg` | Hand touching LogiTemp interface | User interaction, ease-of-use messaging |
| `logitemp-evap-coil.jpg` | Controller mounted on evaporator | Technical/installation context |

### Master-Bilt — Walk-Ins (`assets/products/masterbilt/walk-ins/`)

| File | Description | Best For |
|---|---|---|
| `quick-ship.jpg` | Quick Ship walk-in | Quick Ship promos, speed-to-site messaging |
| `foodservice.jpg` | Foodservice walk-in | General walk-in content |
| `ready-bilt.png` | Ready-Bilt prefabricated walk-in | Ready-Bilt promos, fast install messaging |
| `glass-door.jpg` | Glass door walk-in | Beer cave, retail, visibility messaging |
| `qs-capsule-pak.jpg` | QS Series with Capsule Pak ECO | Self-contained promos for Master-Bilt |
| `bilt2spec-custom.jpg` | Bilt2Spec custom walk-in | Custom/engineered solutions |

### Master-Bilt — Endless Merchandisers (`assets/products/masterbilt/merchandisers/`)

| File | Description | Best For |
|---|---|---|
| `endless-3door.jpg` | BEM/BEL 3-door self-contained | Standard merchandiser hero shot |
| `endless-5door.jpg` | BEM/BEL 5-door self-contained | Large format, high-volume retail |
| `endless-cstore-application.jpg` | Endless in convenience store setting | Application/lifestyle context, retail stories |
| `endless-packed-out.jpg` | 3-door merchandiser packed with product | Capacity/merchandising effectiveness |

### Master-Bilt — Ice Cream Cabinets (`assets/products/masterbilt/ice-cream/`)

| File | Description | Best For |
|---|---|---|
| `dd-66-dipping.jpg` | DD-66 dipping cabinet | Ice cream/frozen dessert, scooping operations |
| `flr-80-floor-display.jpg` | FLR-80 floor display freezer | Floor display, grab-and-go retail |
| `dc-8d-dipping.jpg` | DC-8D dipping cabinet | Smaller dipping operation |

### Master-Bilt — Refrigeration Systems (`assets/products/masterbilt/refrigeration/`)

| File | Description | Best For |
|---|---|---|
| `mrs-front.jpg` | MRS modular refrigeration system | Modular refrigeration, system flexibility |
| `m-series-2fan-covered.jpg` | M-Series 2-fan condensing unit with cover | Remote condensing, mechanical room |
| `split-pak-remote.jpg` | Split-Pak remote refrigeration | Remote refrigeration systems |
| `drs.jpg` | DRS system | Direct refrigeration system |

### Norlake Scientific (`assets/products/norlake-scientific/`)

| File | Description | Best For |
|---|---|---|
| `enviro-line-plasma-freezer.jpg` | Enviro-Line blood plasma freezer | Lab/pharma, blood bank, compliance |
| `mini-room.jpg` | Mini room with control panel | Compact lab storage, research |
| `scipak-system.jpg` | SciPak refrigeration system | Scientific refrigeration systems |

### Outdoor Walk-Ins (`assets/products/outdoor/`)

| File | Description | Best For |
|---|---|---|
| `outdoor-walkin-1.png` | Outdoor walk-in unit | Outdoor applications, weather resistance |
| `outdoor-walkin-2.png` | Outdoor walk-in alternate | Alternate angle |

### Installation Photos (`assets/products/install/`)

| File | Description | Best For |
|---|---|---|
| `install-1.jpg` | Installation in progress | Installation services, setup process |
| `install-2.jpg` | Walk-in installation | Service capability, on-site work |
| `install-3.jpg` | Installation detail | Installation quality, craftsmanship |

### Backgrounds (`assets/backgrounds/`)

| File | Description | Best For |
|---|---|---|
| `mountain-range.jpg` | Mountain landscape | Section dividers, aspirational backdrops |
