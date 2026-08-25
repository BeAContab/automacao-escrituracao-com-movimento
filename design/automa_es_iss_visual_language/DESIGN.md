---
name: Automações ISS Visual Language
colors:
  surface: '#f0fbfd'
  surface-dim: '#d1dcde'
  surface-bright: '#f0fbfd'
  surface-container-lowest: '#ffffff'
  surface-container-low: '#ebf6f8'
  surface-container: '#e5f0f2'
  surface-container-high: '#dfeaec'
  surface-container-highest: '#d9e4e6'
  on-surface: '#131d1f'
  on-surface-variant: '#514537'
  inverse-surface: '#283234'
  inverse-on-surface: '#e8f3f5'
  outline: '#837565'
  outline-variant: '#d5c4b1'
  surface-tint: '#845400'
  primary: '#845400'
  on-primary: '#ffffff'
  primary-container: '#e8a74c'
  on-primary-container: '#623e00'
  inverse-primary: '#fdba5d'
  secondary: '#47626e'
  on-secondary: '#ffffff'
  secondary-container: '#c7e4f2'
  on-secondary-container: '#4b6772'
  tertiary: '#48626b'
  on-tertiary: '#ffffff'
  tertiary-container: '#9db8c2'
  on-tertiary-container: '#2f4952'
  error: '#ba1a1a'
  on-error: '#ffffff'
  error-container: '#ffdad6'
  on-error-container: '#93000a'
  primary-fixed: '#ffddb5'
  primary-fixed-dim: '#fdba5d'
  on-primary-fixed: '#2a1800'
  on-primary-fixed-variant: '#643f00'
  secondary-fixed: '#cae7f4'
  secondary-fixed-dim: '#aecbd8'
  on-secondary-fixed: '#001f28'
  on-secondary-fixed-variant: '#2f4a55'
  tertiary-fixed: '#cbe7f2'
  tertiary-fixed-dim: '#afcbd5'
  on-tertiary-fixed: '#021f27'
  on-tertiary-fixed-variant: '#314b53'
  background: '#f0fbfd'
  on-background: '#131d1f'
  surface-variant: '#d9e4e6'
typography:
  headline-lg:
    fontFamily: Plus Jakarta Sans
    fontSize: 32px
    fontWeight: '700'
    lineHeight: 40px
    letterSpacing: -0.02em
  headline-md:
    fontFamily: Plus Jakarta Sans
    fontSize: 24px
    fontWeight: '700'
    lineHeight: 32px
    letterSpacing: -0.01em
  headline-sm:
    fontFamily: Plus Jakarta Sans
    fontSize: 20px
    fontWeight: '600'
    lineHeight: 28px
  body-lg:
    fontFamily: Hanken Grotesk
    fontSize: 16px
    fontWeight: '400'
    lineHeight: 24px
  body-md:
    fontFamily: Hanken Grotesk
    fontSize: 14px
    fontWeight: '400'
    lineHeight: 20px
  body-sm:
    fontFamily: Hanken Grotesk
    fontSize: 12px
    fontWeight: '400'
    lineHeight: 16px
  label-md:
    fontFamily: Hanken Grotesk
    fontSize: 12px
    fontWeight: '600'
    lineHeight: 16px
    letterSpacing: 0.05em
  headline-lg-mobile:
    fontFamily: Plus Jakarta Sans
    fontSize: 28px
    fontWeight: '700'
    lineHeight: 36px
rounded:
  sm: 0.25rem
  DEFAULT: 0.5rem
  md: 0.75rem
  lg: 1rem
  xl: 1.5rem
  full: 9999px
spacing:
  base: 4px
  xs: 4px
  sm: 8px
  md: 16px
  lg: 24px
  xl: 40px
  gutter: 20px
  margin: 32px
---

## Brand & Style

The design system is engineered for professional accounting automation, prioritizing clarity, trust, and high-efficiency data density. The aesthetic is **Modern Corporate**, blending the precision of fintech with the accessibility of modern SaaS. 

The UI communicates reliability through a grounded petrol blue foundation, while the golden-orange primary color acts as a high-visibility guide for actions and status updates. The interface avoids unnecessary decorative elements, focusing instead on structural integrity, clear information hierarchy, and a systematic approach to complex financial workflows. It is optimized for long-session usage where reduced eye strain and quick scanability are paramount.

## Colors

The palette is anchored by **Dark Petrol Blue (#3B5661)**, used for structural navigation components to provide a sense of stability and institutional trust. **Golden Orange (#E8A74C)** is the strategic primary color, reserved exclusively for primary calls-to-action, active selection states, and critical highlights.

- **Primary**: Used for buttons, active toggle states, and progress indicators.
- **Secondary/Surface**: Dark Petrol Blue is the primary surface color for high-contrast sidebars and footers.
- **Interactive**: Secondary Petrol Blue (#49636C) provides subtle feedback for hover states on dark surfaces.
- **Typography**: All primary text uses Dark Text (#172B31) to ensure WCAG AA compliance against white backgrounds. Neutral Grays are reserved for borders, breadcrumbs, and disabled states.

## Typography

This design system utilizes **Plus Jakarta Sans** for headings to inject a modern, geometric confidence into the interface. For data-heavy environments and body copy, **Hanken Grotesk** is employed for its exceptional legibility at small scales and professional, neutral tone.

- **Headings**: Use Bold weights (700) for primary titles to establish a strong hierarchy.
- **Data Tables**: Use `body-sm` for dense financial grids to maximize information visibility without compromising readability.
- **Labels**: Small caps with increased letter spacing should be used for table headers and section overlines to differentiate them from interactive data points.

## Layout & Spacing

The layout follows a **Fluid Grid** model with a 12-column structure for desktop. To accommodate information-dense accounting workflows, the system uses a compact 4px baseline grid.

- **Margins & Gutters**: Desktop views utilize 32px outer margins with 20px gutters. On mobile, margins reduce to 16px.
- **Density**: Use "Compact" spacing (8px) for data entry forms and "Spacious" spacing (24px) for dashboard overview cards.
- **Alignment**: All financial figures should be right-aligned in tables to facilitate quick vertical comparison. Labels and text descriptions remain left-aligned.

## Elevation & Depth

This design system uses a **Tonal Layering** approach combined with subtle elevation to maintain a clean, flat SaaS aesthetic. 

- **Level 0 (Background)**: Solid #FFFFFF.
- **Level 1 (Cards/Sections)**: Solid #FFFFFF with a 1px border (#B5C0C2) and a very soft, diffused shadow: `0px 4px 12px rgba(23, 43, 49, 0.05)`.
- **Level 2 (Dropdowns/Modals)**: Increased shadow depth to indicate focus: `0px 8px 24px rgba(23, 43, 49, 0.12)`.
- **Interactions**: No elevation change on hover; instead, use a subtle background color shift (e.g., from white to a 5% opacity Petrol Blue) to indicate interactivity.

## Shapes

The shape language strikes a balance between professional rigor and modern approachability. 

- **Standard Elements**: Buttons, Input fields, and Cards utilize a **8px (rounded)** corner radius.
- **Small Elements**: Chips, Tags, and Checkboxes utilize a **4px (soft)** radius to maintain visual weight at smaller scales.
- **Containers**: Large dashboard containers or side panels can scale up to **12px (rounded-lg)** to create a distinct framing effect.

## Components

- **Buttons**: Primary buttons are solid Golden Orange with white text. Secondary buttons use Dark Petrol Blue outlines. Tertiary buttons are text-only with the Petrol Blue color.
- **Input Fields**: 8px radius, 1px #B5C0C2 border. On focus, the border transitions to Dark Petrol Blue with a 2px offset "glow" of the same color at 10% opacity.
- **Data Tables**: High-density rows (40px height). Alternate row striping is not used; instead, use thin 1px horizontal dividers in Neutral Gray. Hover states on rows should use a light tint of Petrol Blue.
- **Status Chips**: Use "pill" shapes (full radius). Success (Green), Warning (Orange), and Error (Red) should use desaturated background tints with high-contrast text for professional legibility.
- **Sidebar**: Fixed width (260px), #3B5661 background. Active nav items use a 4px Golden Orange vertical "indicator" on the left edge and a #49636C background highlight.
- **Cards**: Minimalist white containers with a 1px border. Title areas should be separated by a subtle divider.