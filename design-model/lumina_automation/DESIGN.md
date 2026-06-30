---
name: Lumina Automation
colors:
  surface: '#10131a'
  surface-dim: '#10131a'
  surface-bright: '#363941'
  surface-container-lowest: '#0b0e15'
  surface-container-low: '#191b23'
  surface-container: '#1d2027'
  surface-container-high: '#272a31'
  surface-container-highest: '#32353c'
  on-surface: '#e1e2ec'
  on-surface-variant: '#c2c6d6'
  inverse-surface: '#e1e2ec'
  inverse-on-surface: '#2e3038'
  outline: '#8c909f'
  outline-variant: '#424754'
  surface-tint: '#adc6ff'
  primary: '#adc6ff'
  on-primary: '#002e6a'
  primary-container: '#4d8eff'
  on-primary-container: '#00285d'
  inverse-primary: '#005ac2'
  secondary: '#4cd7f6'
  on-secondary: '#003640'
  secondary-container: '#03b5d3'
  on-secondary-container: '#00424e'
  tertiary: '#ffb786'
  on-tertiary: '#502400'
  tertiary-container: '#df7412'
  on-tertiary-container: '#461f00'
  error: '#ffb4ab'
  on-error: '#690005'
  error-container: '#93000a'
  on-error-container: '#ffdad6'
  primary-fixed: '#d8e2ff'
  primary-fixed-dim: '#adc6ff'
  on-primary-fixed: '#001a42'
  on-primary-fixed-variant: '#004395'
  secondary-fixed: '#acedff'
  secondary-fixed-dim: '#4cd7f6'
  on-secondary-fixed: '#001f26'
  on-secondary-fixed-variant: '#004e5c'
  tertiary-fixed: '#ffdcc6'
  tertiary-fixed-dim: '#ffb786'
  on-tertiary-fixed: '#311400'
  on-tertiary-fixed-variant: '#723600'
  background: '#10131a'
  on-background: '#e1e2ec'
  surface-variant: '#32353c'
typography:
  headline-xl:
    fontFamily: Outfit
    fontSize: 40px
    fontWeight: '700'
    lineHeight: 48px
    letterSpacing: -0.02em
  headline-lg:
    fontFamily: Outfit
    fontSize: 32px
    fontWeight: '600'
    lineHeight: 40px
    letterSpacing: -0.01em
  headline-lg-mobile:
    fontFamily: Outfit
    fontSize: 24px
    fontWeight: '600'
    lineHeight: 32px
  body-md:
    fontFamily: Inter
    fontSize: 16px
    fontWeight: '400'
    lineHeight: 24px
  body-sm:
    fontFamily: Inter
    fontSize: 14px
    fontWeight: '400'
    lineHeight: 20px
  label-caps:
    fontFamily: Inter
    fontSize: 12px
    fontWeight: '600'
    lineHeight: 16px
    letterSpacing: 0.05em
  mono-console:
    fontFamily: JetBrains Mono
    fontSize: 13px
    fontWeight: '400'
    lineHeight: 20px
rounded:
  sm: 0.25rem
  DEFAULT: 0.5rem
  md: 0.75rem
  lg: 1rem
  xl: 1.5rem
  full: 9999px
spacing:
  container-max: 1440px
  gutter: 24px
  margin-mobile: 16px
  stack-sm: 8px
  stack-md: 16px
  stack-lg: 32px
---

## Brand & Style

The design system is engineered for **BeAContab**, a high-performance fiscal automation platform. The brand personality is authoritative yet visionary, combining the precision of accounting with the velocity of modern automation. 

The aesthetic is a sophisticated **Glassmorphism** executed within a deep-space **Dark Mode** environment. It evokes an emotional response of "controlled power"—users should feel they are operating an advanced command center rather than a traditional spreadsheet tool. The interface utilizes translucent layers, precision linework, and vibrant neon accents to highlight active data processing and system health.

## Colors

The palette is anchored in a deep, near-black slate (#0A0C10) to minimize eye strain during long periods of data auditing. 

- **Primary (Cobalt & Neon Blue):** Used for primary actions and brand identity. While #1F4E78 provides a grounded corporate base, #3B82F6 is used for interactive elements to ensure high visibility against the dark background.
- **Success (Neon Cyan):** Specifically reserved for execution states, automated task completion, and "System Healthy" indicators.
- **Accents:** Laranja (#F97316) for warnings and Red (#EF4444) for critical fiscal errors or failed automations.
- **Surface:** Surfaces use a layered glass approach with `backdrop-filter: blur(12px)` to maintain depth and legibility over subtle background gradients.

## Typography

This design system employs a dual-font strategy: **Outfit** for headlines to provide a modern, geometric technological feel, and **Inter** for UI elements and body text to ensure maximum legibility in data-heavy fiscal environments.

A special **Mono** style is introduced for the "Automation Logs" and "Terminal" components, ensuring that technical data is aligned and readable. Use `label-caps` for table headers and small metadata categories to maintain a disciplined, professional hierarchy.

## Layout & Spacing

The layout follows a **Fluid Grid** model with a maximum container width of 1440px. The system relies on a consistent 8px base unit for all padding and margins.

- **Desktop:** 12-column grid with 24px gutters.
- **Sidebar:** Fixed at 280px to accommodate complex fiscal navigation.
- **Data Density:** In spreadsheet views, vertical padding is reduced (compact mode) to allow more rows of data to be visible without scrolling.
- **Responsive:** On mobile, margins shrink to 16px, and complex data tables must transition into card-based layouts.

## Elevation & Depth

Depth is conveyed through **Backdrop Blurs** rather than traditional heavy shadows. 

1. **Base Layer:** The deepest level (#0A0C10) with subtle radial gradients of Cobalt Blue in the corners.
2. **Glass Tier (Cards):** Fundo `rgba(255, 255, 255, 0.05)` with a 1px border of `rgba(255, 255, 255, 0.1)`. This creates a "frosted" look that separates content from the background.
3. **Floating Tier (Modals/Popovers):** Higher opacity glass `rgba(255, 255, 255, 0.08)` with a 12px ambient outer glow in the primary color (#3B82F6) at 10% opacity.

## Shapes

The design system utilizes a **Rounded** corner strategy (8px/0.5rem base) to soften the technical nature of the application and make it feel more approachable. 

- **Standard Buttons & Inputs:** 8px radius.
- **Large Container Cards:** 16px (rounded-lg) to emphasize the "glass pane" metaphor.
- **Progress Bars:** Fully rounded (pill-shaped) to represent fluidity and motion.

## Components

### Buttons
- **Primary:** Gradient background (Cobalt to Neon Blue), white text, subtle outer glow on hover.
- **Ghost:** Glass-style background with 1px border; text in Primary color.

### Input Fields
- **Dark Inputs:** Fundo `rgba(0, 0, 0, 0.2)`, 1px border `rgba(255, 255, 255, 0.1)`. Focus state triggers a Cyan Neon (#06B6D4) border glow.

### Console Log (Terminal)
- A specialized component for real-time automation feedback. Darker background than standard cards, monospaced font, and syntax highlighting for success/error messages.

### Progress Bars
- Ultra-thin (4px height). Active state features a "pulse" animation traveling through the bar to signify active processing.

### Chips/Badges
- Small, rounded elements with high-saturation backgrounds at low opacity (e.g., Success badge has a background of Cyan at 15% opacity with 100% opacity Cyan text).

### Animations
- **Processing:** A subtle breathing pulse on the Primary button during automation execution.
- **Transitions:** Horizontal slide for tab switching and 200ms fade-in for glass cards.