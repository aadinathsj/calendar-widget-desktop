# Calendar Widget UI – Premium Glassmorphism Redesign

## Overview
Complete UI redesign of the Outlook Calendar Widget with Apple-inspired glassmorphism aesthetic, optimized for Windows 11 dark navy + electric blue wallpapers. Built with pure HTML/CSS/Vanilla JS – zero frameworks, zero emojis, production-ready.

## Design Language

### Visual Style
- **Liquid Glass Surface**: Translucent frosted panels with backdrop-filter blur (24px), layered depth via nested materials
- **Continuous Curvature**: 24px primary radius, 16px cards, 12px buttons – creates soft, organic feel
- **Layered Shadows**: Multiple soft shadows (8-64px blur) with low opacity for realistic elevation
- **Subtle Borders**: Dual-stroke system (outer 14% + inner 8% white) mimics glass refraction
- **Specular Highlight**: Gradient overlay on main container suggests light reflection

### Color System
Tuned for dark navy wallpaper with electric blue accents:
- **Base**: `rgba(11, 18, 32, 0.45)` – deep navy with transparency
- **Accents**: `#3B82F6` (primary blue), `#5BCBFF` (cyan highlight)
- **Text**: Near-white `#EAF0FF` primary, desaturated `#B7C6E6` secondary
- **Status**: Blue (upcoming), cyan glow (active), muted gray (past)

### Typography
- **Family**: Segoe UI Variable → Segoe UI → system-ui fallback
- **Hierarchy**: 15px semibold headers, 13px medium meeting titles, 11-13px metadata
- **Line-height**: 1.3 for headers, 1.4-1.5 for body – tight but readable

### Spacing
8px grid system throughout:
- `--space-1: 8px` to `--space-6: 48px`
- Consistent rhythm in padding, gaps, margins

## Layout Structure

### 1. Fixed Header (56px)
```
┌─────────────────────────────────────────┐
│ Today · 3 events      [ ‹ › ↻ ]         │
└─────────────────────────────────────────┘
```
- Left: Date label + event count
- Right: Compact pill group with prev/next/refresh controls
- Elevated glass material (darker than base)

### 2. Scrollable Timeline (flex: 1)
```
┌─────────────────────────────────────────┐
│ ░░░░░░░░ (fade mask)                    │
│ ┌─────────────────────────────────────┐ │
│ │ ▌9:00 AM  Team Standup              │ │
│ │           📍 Conference Room A       │ │
│ │           [Join] [Notes]             │ │
│ └─────────────────────────────────────┘ │
│ ┌─────────────────────────────────────┐ │
│ │ ▌10:00 AM Project Review...         │ │
│ ░░░░░░░░ (fade mask)                    │
└─────────────────────────────────────────┘
        │
        └─ Custom scrollbar (6px, cyan)
```

**Scroll Features**:
- Custom webkit scrollbar (6px width, rounded, cyan thumb)
- Top/bottom fade masks (32px gradient) appear when scrollable
- Smooth scroll behavior
- Auto-scroll to selected meeting

**Meeting Cards**:
- Left accent bar (3px) indicates status: blue/cyan/gray
- Hover: lift 1px, brighten background, increase border opacity
- Selected: accent border + glow shadow
- Focus-visible: 2px blue outline

### 3. Notes Sheet Overlay (slide-up)
```
┌─────────────────────────────────────────┐
│ Project Review       10:00-11:30   [×]  │
├─────────────────────────────────────────┤
│                                          │
│ [scrollable textarea]                    │
│                                          │
├─────────────────────────────────────────┤
│ Saved                          [→]      │
└─────────────────────────────────────────┘
```
- Slides up from bottom (300ms ease-out)
- Same glass material as main container
- Scrollable textarea with custom scrollbar
- Auto-save (800ms debounce) with "Saving..." → "Saved" indicator
- Esc key to close

## Interactions

### Hover States
- Duration: 160-220ms ease-out
- Effects: background brightness +10%, border opacity +6%, translateY(-1px)
- Buttons: subtle scale or lift

### Focus States
- All interactive elements: 2px solid accent outline, 2px offset
- Visible keyboard navigation
- Logical tab order

### Keyboard Navigation
| Key | Action |
|-----|--------|
| `Arrow Up/Down` | Navigate meetings (wraps) |
| `Arrow Left/Right` | Previous/next day |
| `Enter` | Open notes for selected meeting |
| `Esc` | Close notes |
| `Tab` | Standard focus order |

### Animations
- Notes sheet: `transform: translateY()` + opacity fade
- Refresh button: 360° rotation on click
- Meeting selection: smooth scroll to view
- All transitions use `cubic-bezier(0.4, 0, 0.2, 1)`

## Technical Implementation

### CSS Architecture
1. **Design Tokens** (lines 1-60): All colors, spacing, radii as CSS variables
2. **Base Styles** (61-85): Reset, body, font smoothing
3. **Glassmorphism** (86-120): Main container with backdrop-filter
4. **Components** (121+): Header, timeline, cards, notes – modular structure

### JavaScript Features
- **State**: `currentDate`, `selectedMeetingIndex`, `meetings` array
- **Rendering**: Pure DOM manipulation (no template library)
- **Mock Data**: `generateMockMeetings()` for demo (replace with Electron API)
- **Auto-save**: Debounced localStorage save (production: electron-store)
- **Accessibility**: ARIA labels, roles, keyboard nav, focus management

### Integration Points (Electron)
Replace mock implementations:
```javascript
// app.js lines to connect:
loadMeetingsForDate() → window.electronAPI.getEvents(start, end)
saveNotes() → window.electronAPI.saveNote(id, content, metadata)
openFileLocation() → window.electronAPI.openNotesFile(id)
openTeamsLink() → window.electronAPI.openExternal(url)
```

## Accessibility

### WCAG AA Compliance
- **Contrast**: Text colors meet 4.5:1 minimum on glass backgrounds
- **Focus**: All interactive elements have visible focus indicators
- **Keyboard**: Full navigation without mouse
- **ARIA**: Proper labels, roles, hidden states, selected states
- **Reduced Motion**: No animation preference respected (add if needed)

### Prefers Reduced Transparency
```css
@media (prefers-reduced-transparency: reduce) {
  --glass-base-bg: rgba(11, 18, 32, 0.92); /* increase opacity */
  --glass-base-blur: 8px; /* reduce blur */
}
```

## Responsive Considerations

### Windows Scaling
- Designed at 360×520px (100% scale)
- Tested concepts: 125%, 150% scale
- Uses relative units (px converted to rem if needed)
- System fonts scale naturally

### Contrast Adjustments
For different wallpapers, tune:
```css
--glass-base-bg: /* opacity 0.45-0.75 */
--text-primary: /* brightness adjust */
--glass-border: /* opacity 0.14-0.24 */
```

## File Structure
```
src/renderer/
├── index.html          # Semantic structure, meeting cards, notes sheet
├── styles.css          # Design tokens, glassmorphism, components
└── app.js              # State, rendering, keyboard nav, auto-save
```

## Performance Notes
- **Backdrop-filter**: GPU-accelerated, smooth on modern hardware
- **Transitions**: Only transform/opacity (no layout thrash)
- **Scrolling**: Optimized with `scroll-behavior: smooth`
- **Rendering**: Efficient DOM updates (no full re-renders)

## Browser/Electron Support
- Chrome/Edge 76+ (backdrop-filter)
- Electron 6+ (modern Chromium)
- Fallback: Increase opacity if backdrop-filter unsupported

## Future Enhancements
1. **Dark/Light mode toggle** – swap color tokens
2. **Density settings** – compact/comfortable/spacious spacing
3. **Custom accent colors** – user preference
4. **Markdown support** – rich text in notes editor
5. **Drag-to-resize** – adjustable widget dimensions

---

**Design Philosophy**: Calm, precise, premium. Every pixel serves clarity. No decoration without purpose. Glassmorphism as structure, not ornament.
