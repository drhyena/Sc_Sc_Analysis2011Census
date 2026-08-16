import asyncio
import re
from pathlib import Path

from playwright.async_api import async_playwright


# ============================================================
# EXACT URL FROM YOUR SCREENSHOT
# ============================================================

URL = "https://censusindia.gov.in/census.website/data/census-tables#"

OUTPUT_DIR = Path(r"C:\Users\PC\Desktop\ScSctdata\census_SC_ST")


# ============================================================
# RESUME POINT -- no longer hardcoded. These get filled in at
# runtime by wait_for_click_to_resume(), based on the EXCEL
# button you click in the live browser window before automation
# starts. Leave them as None here.
# ============================================================

RESUME_TABLE_KEYWORD = None
RESUME_STATE = None

# Once both the resume card AND resume state have been found (or
# if you chose to start from the very beginning), every card/state
# after that point downloads normally -- no more skipping.
# This flag is mutated via `global` inside the functions below.
resumed = False


# ============================================================
# HELPERS
# ============================================================

def clean_filename(name):
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    name = re.sub(r"\s+", " ", name)
    return name.strip()[:180]


async def get_card_from_tag(tag, tag_text):
    """
    Starting from the Scheduled Castes / Scheduled Tribes
    tag, walk UP the DOM until we find the table card.

    We specifically look for an ancestor containing:

        Scheduled Castes / Scheduled Tribes
        tables available

    """

    for level in range(1, 12):

        try:

            xpath = "/.." * level

            candidate = tag.locator(
                "xpath=" + xpath
            )

            text = await candidate.inner_text()

            lower = text.lower()

            if (
                "tables available" in lower
                and tag_text.lower() in lower
                and len(text) > 100
            ):

                return candidate

        except Exception:
            pass

    return None


async def find_sc_st_cards(page):
    """
    Find actual cards containing the visible tags:

        Scheduled Castes
        Scheduled Tribes
    """

    cards = []

    for tag_text in [
        "Scheduled Castes",
        "Scheduled Tribes",
    ]:

        locator = page.get_by_text(
            tag_text,
            exact=True
        )

        count = await locator.count()

        print(
            f"  {tag_text}: {count} tag(s)"
        )

        for i in range(count):

            tag = locator.nth(i)

            card = await get_card_from_tag(
                tag,
                tag_text
            )

            if card is not None:

                try:

                    card_text = (
                        await card.inner_text()
                    )

                    # First meaningful line is normally
                    # the table title.
                    lines = [
                        x.strip()
                        for x in card_text.splitlines()
                        if x.strip()
                    ]

                    title = (
                        lines[0]
                        if lines
                        else tag_text
                    )

                    cards.append(
                        {
                            "card": card,
                            "tag": tag_text,
                            "title": title,
                            "full_text": card_text,
                        }
                    )

                except Exception:
                    pass

    return cards


async def expand_card(card):
    """
    Click the '31 tables available' /
    '30 tables available' dropdown.
    """

    # Look for text such as:
    #
    # 31 tables available
    # 30 tables available

    control = card.get_by_text(
        re.compile(
            r"\d+\s+tables?\s+available",
            re.I
        )
    )

    count = await control.count()

    if count == 0:
        return False

    try:

        await control.first.scroll_into_view_if_needed()

        await control.first.click(
            timeout=5000
        )

        await asyncio.sleep(1)

        return True

    except Exception:

        return False


async def get_row_state_name(element):
    """
    Given an EXCEL button element, find its row and pull out the
    STATE/DISTRICT label from that row. Tries a real <tr> ancestor
    first (most likely, given the tabular layout); falls back to
    walking up generic ancestors and taking the shortest one that
    still contains 'EXCEL' -- that's the row, not the whole card.
    """

    # Try a real table row first.
    try:

        tr = element.locator("xpath=ancestor::tr[1]")

        if await tr.count() > 0:

            text = await tr.first.inner_text()

            lines = [
                x.strip()
                for x in text.splitlines()
                if x.strip()
            ]

            if lines:
                return lines[0]

    except Exception:
        pass

    # Fallback: walk up generic ancestors, take the smallest one
    # that still contains "EXCEL" -- that's the row-level wrapper.
    for level in range(1, 8):

        try:

            xpath = "/.." * level

            candidate = element.locator(
                "xpath=" + xpath
            )

            text = await candidate.inner_text()

            if "excel" in text.lower() and len(text) < 200:

                lines = [
                    x.strip()
                    for x in text.splitlines()
                    if x.strip()
                ]

                if lines:
                    return lines[0]

        except Exception:
            pass

    return None


async def find_excel_buttons(card):
    """
    Find EXCEL links/buttons belonging ONLY to this card, paired
    with the STATE/DISTRICT row name each one belongs to.

    Returns a list of dicts: {"state_name": str or None, "button": element}
    """

    elements = card.locator(
        "a, button"
    )

    count = await elements.count()

    results = []

    for i in range(count):

        element = elements.nth(i)

        try:

            text = (
                await element.inner_text()
            ).strip()

            aria = (
                await element.get_attribute(
                    "aria-label"
                )
                or ""
            )

            title = (
                await element.get_attribute(
                    "title"
                )
                or ""
            )

            href = (
                await element.get_attribute(
                    "href"
                )
                or ""
            )

            combined = (
                text
                + " "
                + aria
                + " "
                + title
                + " "
                + href
            ).lower()

            if (
                "excel" in combined
                or ".xls" in combined
                or ".xlsx" in combined
            ):

                state_name = await get_row_state_name(element)

                results.append(
                    {
                        "state_name": state_name,
                        "button": element,
                    }
                )

        except Exception:
            pass

    return results

async def download_excel(page, button, table_title, state_number):

    try:
        async with page.expect_download(timeout=60000) as info:
            await button.click()

        download = await info.value

        filename = download.suggested_filename

        if not filename:
            filename = f"{clean_filename(table_title)}_{state_number}.xlsx"

        filename = clean_filename(filename)

        destination = OUTPUT_DIR / filename

        if destination.exists():
            stem = destination.stem
            suffix = destination.suffix

            n = 2

            while destination.exists():
                destination = OUTPUT_DIR / f"{stem}_{n}{suffix}"
                n += 1

        await download.save_as(str(destination))

        print(f"        ✓ {filename}")

        return True

    except Exception as e:

        print(f"        ✗ {e}")

        return False


# ============================================================
# CLICK-TO-RESUME
#
# Instead of hardcoding which table/state to resume from, we let
# the user drive the live (headless=False) browser: navigate
# pagination, expand whatever card they want, and click the exact
# EXCEL button they want to start at. That click is intercepted
# (prevented from actually downloading) and used to identify the
# table title + state/district row, which become the resume point
# for the normal automated flow below.
# ============================================================

async def wait_for_click_to_resume(page):
    """
    Pause here and let the user click the EXCEL button they want to
    resume from. Returns (card_title, state_name). Returns
    (None, None) if the user presses Escape (or nothing is clicked
    within the timeout) to signal "start from the very beginning".
    """

    print()
    print("=" * 70)
    print("CLICK THE EXCEL BUTTON YOU WANT TO START FROM")
    print("Navigate pagination / expand a table card in the browser")
    print("window, then click the EXCEL link for the row you want to")
    print("resume at. The click will NOT actually download anything --")
    print("it just marks the starting point. Automation takes over")
    print("right after that.")
    print()
    print("Press Escape on the page instead to start from the very")
    print("first table. (Auto-starts from the beginning after 10")
    print("minutes of no click.)")
    print("=" * 70)
    print()

    result = {"picked": None}
    done = asyncio.Event()

    async def on_pick(marker_id):
        result["picked"] = marker_id
        done.set()

    async def on_skip():
        done.set()

    await page.expose_function("__resumeClickPicked", on_pick)
    await page.expose_function("__resumeClickSkip", on_skip)

    await page.evaluate(
        """
        () => {
            const banner = document.createElement('div');
            banner.textContent =
                'Click the EXCEL button to resume from  (Esc = start from the beginning)';
            banner.id = '__resume_banner__';
            banner.style.position = 'fixed';
            banner.style.top = '0';
            banner.style.left = '0';
            banner.style.right = '0';
            banner.style.zIndex = '2147483647';
            banner.style.background = '#ffdd00';
            banner.style.color = '#000';
            banner.style.fontSize = '16px';
            banner.style.fontWeight = 'bold';
            banner.style.padding = '10px';
            banner.style.textAlign = 'center';
            document.body.appendChild(banner);

            const isExcelish = (el) => {
                if (!el) return false;
                const text = (el.innerText || '').toLowerCase();
                const href = (el.getAttribute('href') || '').toLowerCase();
                const aria = (el.getAttribute('aria-label') || '').toLowerCase();
                const title = (el.getAttribute('title') || '').toLowerCase();
                const combined = text + ' ' + href + ' ' + aria + ' ' + title;
                return combined.includes('excel') || combined.includes('.xls');
            };

            const clickHandler = (e) => {
                let el = e.target;
                while (el && el !== document.body) {
                    if ((el.tagName === 'A' || el.tagName === 'BUTTON') && isExcelish(el)) {
                        e.preventDefault();
                        e.stopPropagation();

                        let markerId = el.getAttribute('data-resume-marker');
                        if (!markerId) {
                            markerId = 'marker-' + Math.random().toString(36).slice(2);
                            el.setAttribute('data-resume-marker', markerId);
                        }

                        el.style.outline = '4px solid red';
                        window.__resumeClickPicked(markerId);
                        document.removeEventListener('click', clickHandler, true);
                        document.removeEventListener('keydown', keyHandler, true);
                        banner.remove();
                        return;
                    }
                    el = el.parentElement;
                }
            };

            const keyHandler = (e) => {
                if (e.key === 'Escape') {
                    window.__resumeClickSkip();
                    document.removeEventListener('click', clickHandler, true);
                    document.removeEventListener('keydown', keyHandler, true);
                    banner.remove();
                }
            };

            document.addEventListener('click', clickHandler, true);
            document.addEventListener('keydown', keyHandler, true);
        }
        """
    )

    try:
        await asyncio.wait_for(done.wait(), timeout=600)
    except asyncio.TimeoutError:
        print("No click detected within 10 minutes -- starting from the beginning.")
        return None, None

    marker_id = result["picked"]

    if marker_id is None:
        print("Starting from the very first table.")
        return None, None

    element = page.locator(f'[data-resume-marker="{marker_id}"]')

    state_name = await get_row_state_name(element)

    card_title = None

    for level in range(1, 12):
        try:
            xpath = "/.." * level
            candidate = element.locator("xpath=" + xpath)
            text = await candidate.inner_text()
            lower = text.lower()
            if "tables available" in lower and len(text) > 100:
                lines = [x.strip() for x in text.splitlines() if x.strip()]
                if lines:
                    card_title = lines[0]
                break
        except Exception:
            pass

    print()
    print("Resume point captured:")
    print(f"  Table : {card_title}")
    print(f"  State : {state_name}")
    print()

    return card_title, state_name


async def process_current_page(page):

    """
    Process the current pagination page.
    """

    global resumed

    print()
    print(
        "=" * 70
    )

    print(
        f"URL: {page.url}"
    )

    print(
        "=" * 70
    )

    cards = await find_sc_st_cards(
        page
    )

    print()
    print(
        f"SC/ST cards found: {len(cards)}"
    )

    total_downloaded = 0

    for number, item in enumerate(
        cards,
        start=1
    ):

        card = item["card"]
        category = item["tag"]
        title = item["title"]
        full_text = item["full_text"]

        # ------------------------------------------------------
        # RESUME LOGIC -- skip cards entirely until we hit the
        # resume table, unless we've already resumed.
        # ------------------------------------------------------

        if not resumed:

            card_matches = (
                RESUME_TABLE_KEYWORD is not None
                and (
                    RESUME_TABLE_KEYWORD.lower() in title.lower()
                    or RESUME_TABLE_KEYWORD.lower() in full_text.lower()
                )
            )

            if not card_matches:

                print()
                print(f"[skip] {title}  (before resume point)")

                continue

        print()
        print(
            "-" * 70
        )

        print(
            f"[{number}] {title}"
        )

        print(
            f"     TAG: {category}"
        )

        # ----------------------------------------------------
        # Expand the state table.
        # ----------------------------------------------------

        expanded = await expand_card(
            card
        )

        print(
            f"     Expanded: {expanded}"
        )

        await asyncio.sleep(0.5)

        # ----------------------------------------------------
        # Find EXCEL buttons (each paired with its state name).
        # ----------------------------------------------------

        button_entries = await find_excel_buttons(
            card
        )

        print(
            f"     Excel buttons: "
            f"{len(button_entries)}"
        )

        # ----------------------------------------------------
        # RESUME LOGIC -- within the resume card, skip states
        # before RESUME_STATE. Once found (or if this isn't the
        # resume card, or RESUME_STATE is None), download
        # everything from here on and flip `resumed` on.
        # ----------------------------------------------------

        skipping_states = (not resumed) and (RESUME_STATE is not None)

        # ----------------------------------------------------
        # Download every state Excel (respecting resume skip).
        # ----------------------------------------------------

        for state_number, entry in enumerate(
            button_entries,
            start=1
        ):

            button = entry["button"]
            state_name = entry["state_name"] or ""

            if skipping_states:

                is_target = (
                    state_name.strip().lower()
                    == RESUME_STATE.strip().lower()
                )

                if not is_target:

                    print(
                        f"        [skip] {state_name or '(unknown row)'}"
                    )

                    continue

                # Found the target state -- stop skipping from here on.
                skipping_states = False
                resumed = True

            elif not resumed:
                # Resume card found, no RESUME_STATE given (or already
                # passed it) -- everything downloads, mark resumed.
                resumed = True

            success = await download_excel(
                page,
                button,
                title,
                state_number
            )

            if success:
                total_downloaded += 1

            await asyncio.sleep(0.2)

    return total_downloaded


async def click_next(page):
    """
    Click the actual pagination Next button.
    """

    # The screenshot shows:
    #
    # << First
    # <- Previous
    # Next ->
    # Last >>

    candidates = page.locator(
        "a, button"
    )

    count = await candidates.count()

    for i in range(count):

        element = candidates.nth(i)

        try:

            text = (
                await element.inner_text()
            ).strip().lower()

            # Don't accidentally select "Next" text elsewhere.
            if not (
                text == "next"
                or text == "next →"
                or text == "next →"
                or text.startswith("next")
            ):
                continue

            # Check disabled.
            disabled = (
                await element.get_attribute(
                    "disabled"
                )
            )

            aria_disabled = (
                await element.get_attribute(
                    "aria-disabled"
                )
            )

            classes = (
                await element.get_attribute(
                    "class"
                )
                or ""
            ).lower()

            if disabled is not None:
                continue

            if aria_disabled == "true":
                continue

            if "disabled" in classes:
                continue

            print(
                "Clicking Next..."
            )

            await element.scroll_into_view_if_needed()

            await element.click()

            await asyncio.sleep(2)

            return True

        except Exception:
            pass

    return False


async def main():

    global RESUME_TABLE_KEYWORD, RESUME_STATE, resumed

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    async with async_playwright() as p:

        browser = await p.chromium.launch(
            headless=False
        )

        page = await browser.new_page(
            viewport={
                "width": 1400,
                "height": 1000
            }
        )

        print()
        print(
            "=" * 70
        )

        print(
            "CENSUS INDIA SC/ST EXCEL DOWNLOADER"
        )

        print(
            "=" * 70
        )

        print()
        print(
            "Opening EXACT page:"
        )

        print(
            URL
        )

        # ----------------------------------------------------
        # IMPORTANT: EXACT URL FROM YOUR SCREENSHOT.
        # ----------------------------------------------------

        await page.goto(
            URL,
            wait_until="domcontentloaded",
            timeout=120000
        )

        # This site loads the table catalogue dynamically.
        await asyncio.sleep(8)

        # ------------------------------------------------------
        # Let the user click where to resume from, instead of a
        # hardcoded keyword/state.
        # ------------------------------------------------------

        RESUME_TABLE_KEYWORD, RESUME_STATE = await wait_for_click_to_resume(
            page
        )

        if RESUME_TABLE_KEYWORD is None:
            # No resume point picked -- download everything from
            # the very first card.
            resumed = True

        total_downloaded = 0
        page_number = 1

        while True:

            print()
            print()
            print(
                "#" * 70
            )

            print(
                f"CATALOGUE PAGE {page_number}"
            )

            print(
                "#" * 70
            )

            downloaded = (
                await process_current_page(
                    page
                )
            )

            total_downloaded += downloaded

            print()
            print(
                f"Downloaded on page "
                f"{page_number}: "
                f"{downloaded}"
            )

            # ------------------------------------------------
            # Next catalogue page.
            # ------------------------------------------------

            moved = await click_next(
                page
            )

            if not moved:

                print()
                print(
                    "No more catalogue pages."
                )

                break

            page_number += 1

            # The site has 238 tables / 15 per page,
            # so normally about 16 pages.
            if page_number > 100:

                print(
                    "Safety limit reached."
                )

                break

        await browser.close()

    print()
    print(
        "=" * 70
    )

    print(
        "DONE"
    )

    print(
        "=" * 70
    )

    print(
        f"Total Excel files: "
        f"{total_downloaded}"
    )

    print()
    print(
        "Saved to:"
    )

    print(
        OUTPUT_DIR.resolve()
    )


if __name__ == "__main__":

    asyncio.run(main())