# Design System

This calendar widget follows modern product UI design principles based on the impeccable framework.

## Design Philosophy

**Register**: Product UI - the interface serves the task (viewing calendar, taking notes)  
**Strategy**: Restrained color with tinted neutrals + single accent  
**Target**: Professional, trustworthy, efficient

## Color System

### OKLCH Color Space
Using perceptually uniform OKLCH instead of HSL for consistent lightness across hues.

### Palette Structure

**Tinted Neutrals** (blue hue 250, low chroma 0.003-0.008)
- Subtle blue tint creates cohesion with primary accent
- Never pure gray - adds warmth and polish
- 10-shade scale from near-black to near-white

**Primary Accent** (Microsoft blue, OKLCH)
- Used sparingly: CTAs, selection states, focus indicators
- ~10% visual weight (restrained strategy)
- Three shades: 400 (hover), 500 (default), 600 (active)

**Semantic Colors**
- Success: Green (hue 145)
- Error: Red (hue 25)
- Warning: Amber (hue 75)

### Color Strategy: Restrained

- 60% neutral backgrounds and surfaces
- 30% text and borders
- 10% accent color for primary actions only

## Typography

**Font Stack**: `'Segoe UI', -apple-system, system-ui, sans-serif`
- System fonts for native Windows 11 feel
- Excellent readability at small sizes
- Consistent with Windows design language

**Scale**: Tight ratio (1.125) for product UI
- xs: 0.75rem (labels, timestamps)
- sm: 0.875rem (body, secondary text)
- base: 1rem (primary content)
- lg: 1.125rem (subheadings)
- xl: 1.265rem (headings)
- 2xl: 1.424rem (page titles)

**Hierarchy**
- Weight contrast: 400 (regular), 500 (medium), 600 (semibold)
- Letterspacing: -0.01em for headings, 0.01em for buttons
- Line height: 1.5 for body, 1.4 for headings

## Layout

**Spacing Scale**: 0.25rem base unit
- Consistent rhythm through multiples
- Varied spacing prevents monotony
- Tighter in dense UI areas (event cards)
- More generous for focus areas (notes editor)

**Structure**
- Familiar patterns: top bar, content area, detail overlay
- No unnecessary containers
- Content-first, minimal chrome

## Components

### Event Cards
- Subtle border + hover state (no side-stripe!)
- 3px left accent appears on hover only
- Rounded corners (10px) match Windows 11
- Subtle shadow on hover for depth
- Transform feedback on interaction

### Buttons
- Primary: Accent color, bold weight
- States: default, hover (darker), active (scale), disabled
- Rounded (8px) for approachability
- Clear focus rings for accessibility

### Notes Editor
- Full-height textarea
- Minimalist chrome
- Metadata panel for context
- Slide-in animation (200ms ease-out)

## Motion

**Timing**: 150-250ms range
- 150ms: Fast feedback (buttons, hovers)
- 200ms: Standard transitions (cards, overlays)
- No orchestrated sequences - users are in flow

**Easing**: ease-out only
- Natural deceleration
- No bounce or elastic (product ban)

**Purpose**: State changes only
- Hover states
- Active feedback
- Panel transitions
- No decorative motion

## Shadows

**Layered, subtle elevations**
- sm: Hover states (1px + 2px blur)
- md: Elevated cards (2px + 8px blur)
- lg: Window/modal (2px + 8px, 8px + 24px)
- All use OKLCH black with low opacity

## Accessibility

**Contrast**: WCAG AA minimum
- Body text: 4.5:1
- Large text: 3:1
- UI components: 3:1

**Focus States**: 2px solid outline, 2px offset
**Keyboard Navigation**: All interactive elements focusable
**Screen Readers**: Semantic HTML structure

## What We Avoid

Based on impeccable product bans:

- ❌ Side-stripe borders as decoration
- ❌ Gradient text
- ❌ Display fonts in UI labels
- ❌ Decorative motion
- ❌ Pure black (#000) or pure white (#fff)
- ❌ Inconsistent component vocabulary
- ❌ Reinvented standard affordances

## Windows 11 Integration

**Visual Language**
- Rounded corners (8-12px)
- Subtle shadows and elevation
- Restrained color palette
- System font (Segoe UI)

**Behavior**
- Frameless window with custom controls
- Always-on-top for quick reference
- Top-right corner positioning
- Respects Windows theme (light mode)

## Future Enhancements

Potential additions maintaining the design system:

- Dark mode variant (with adjusted OKLCH lightness values)
- Skeleton loading states
- Empty state illustrations
- Keyboard shortcuts overlay
- Settings panel for customization
