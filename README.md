# SvS Prep Buff Sign-Up — State 2679

Unofficial scheduling tool for Whiteout Survival, State 2679.
Lets alliance members sign up for SvS Prep Buff slots (Construction, Research, Troop Training) on a shared calendar.

## Live page
https://VOTRE-PSEUDO-GITHUB.github.io/NOM-DU-DEPOT/

## How it works
- Players pick their alliance, enter their pseudo and ID, then choose an open 30-minute slot.
- Each day allows one slot per player.
- Sunday's only slot (23:45) feeds into Monday's Construction buff.
- Wednesday's only slot (23:45) feeds into Thursday's Troop Training buff.
- No buff is scheduled on Wednesday or Friday outside of that one carry-over slot.

## Admin panel
Accessible via the small "Admin" link at the bottom of the page. Lets the organizer view, add, or remove sign-ups, merged into 3 views: Monday (Construction), Tuesday (Research), and Thursday (Troop Training).

## Tech notes
Single static HTML file (`index.html`), no build step. Data is stored via Claude's artifact storage API (`window.storage`), shared across all visitors. This means the live page only works correctly if reopened through the original Claude artifact link, not as a plain GitHub Pages deployment, since `window.storage` is specific to that environment.

## Disclaimer
Unofficial community tool. Not affiliated with or endorsed by the developers of Whiteout Survival.v
