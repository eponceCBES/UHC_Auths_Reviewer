"""
WellSky Aging & Disability — reusable automation library.

The WellSky SAMS app is an OpenSilver (Silverlight-port) web app embedded
inside an Angular shell via iframe. Two critical facts drive every selector:

1. The OUTER Angular shell uses normal semantic HTML (inputs, buttons) —
   ordinary CSS selectors work fine there. Key selectors:
     - Global search: input[name='searchControl']   (placeholder 'Search...')
     - Consumer result option: mat-option           (Angular Material)
     - Top-nav buttons by name: "My Dashboard", "Consumers", + overflow items
       (Calls, Routes, Activities, Rosters, Reports, Administrator, Contracts,
       Invoices, Payments, Saved Searches, Tools, Claims, Menu Settings).

2. The INNER iframe (OpenSilver) generates DOM with sequential numeric IDs
   like `id11844`, `id11845` that CHANGE EVERY SESSION. NEVER rely on them.
   Instead, WellSky XAML templates expose a stable hook:

       <div data-id="control-id_<PropertyName>" ...>

   For every bound control. That is the anchor you want. Examples seen:
     control-id_Subject, control-id_ActionUuid, control-id_AgencyUuid,
     control-id_LevelCareLocusCareUuid, control-id_DueDate, control-id_ProviderUuid

   For buttons, the stable pattern is:
       //span[normalize-space()='<ButtonLabel>']

   For tabs, it's text-content matching on the sidebar:
       //span[contains(text(),'<TabName>')]

The helpers below encode these patterns once so higher-level scripts stop
re-implementing them.

Usage:
    from wellsky import WellSkyClient

    w = WellSkyClient(username="<user>", password="<password>")
    w.login()
    w.open_consumer("<consumer-id>")
    w.goto_tab("Activities", "Referrals")
    w.click_add_new()
    w.fill_textarea("Subject", "Follow-up call")
    w.pick_dropdown("ActionUuid", "Reassessment")
    w.pick_dropdown("AgencyUuid", "Central Boston Elder Services, Inc.")
    w.set_date("DueDate", "06/15/2026")
    w.save()
    w.close()
"""

from __future__ import annotations

import os
import time
from dataclasses import dataclass, field
from typing import Iterable, Optional

from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    NoSuchElementException,
    TimeoutException,
    StaleElementReferenceException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


# ────────────────────────────────────────────────────────────────────
# CONSTANTS
# ────────────────────────────────────────────────────────────────────

OKTA_URL = "https://wellsky-ciam.okta.com/"
AGING_URL = "https://agingprod.wellsky.com/aging/prod/#/"

# Known CBES constants — for this user, Agency is always CBES.
DEFAULT_AGENCY = "Central Boston Elder Services, Inc."

# Full dropdown option lists live in `wellsky_options.py` — 8 lists
# harvested from the sandbox (203 Actions, 252 Care Programs, 251
# Providers, 30 Agencies, 42 Sites, 10 Statuses, 3 Reasons, 3
# Follow-up statuses). Import from there when you need to validate a
# value or autocomplete a pick list.
FOLLOWUP_STATUSES = ["Completed", "Not Required", "Required"]

# Outer-shell (Angular) selectors — stable.
SEL_GLOBAL_SEARCH = (By.CSS_SELECTOR, "input[name='searchControl']")
SEL_MAT_OPTION = (By.CSS_SELECTOR, "mat-option")
SEL_OUTER_IFRAME = (By.CSS_SELECTOR, "iframe")

# Top-nav buttons (outer shell).
TOPNAV_BUTTONS = {
    "My Dashboard": "//button[.//*[normalize-space()='My Dashboard']]",
    "Consumers":    "//button[.//*[normalize-space()='Consumers']]",
    "Calls":        "//button[.//*[normalize-space()='Calls']]",
    "Routes":       "//button[.//*[normalize-space()='Routes']]",
    "Activities":   "//button[.//*[normalize-space()='Activities']]",
    "Rosters":      "//button[.//*[normalize-space()='Rosters']]",
    "Reports":      "//button[.//*[normalize-space()='Reports']]",
    "Administrator":"//button[.//*[normalize-space()='Administrator']]",
    "Contracts":    "//menuitem[normalize-space()='Contracts'] | //*[@role='menuitem'][normalize-space()='Contracts']",
    "Invoices":     "//*[@role='menuitem'][normalize-space()='Invoices']",
    "Payments":     "//*[@role='menuitem'][normalize-space()='Payments']",
    "Saved Searches":"//*[@role='menuitem'][normalize-space()='Saved Searches']",
    "Tools":        "//*[@role='menuitem'][normalize-space()='Tools']",
    "Claims":       "//*[@role='menuitem'][normalize-space()='Claims']",
}

# Real consumer sidebar tabs (harvested 2026-04-24 from a CCA sandbox
# consumer). These are the exact span labels — pass them to goto_tab().
CONSUMER_TABS = [
    "Details",
    "Activities & Referrals",
    "Assessments",
    "Billing",
    "Calls",
    "Care Plans",
    "File Attachments",
    "Journals",
    "Routes",
    "Service Deliveries",
    "Service Orders",
]

# Top-of-record toolbar buttons (next to Save / Save and Close).
CONSUMER_TOOLBAR_BUTTONS = [
    "Save", "Save and Close", "Close", "Reject Changes",
    "Print", "Open Audits", "Format Panels", "Status Wizard",
    "Merge", "Copy Client ID", "Add New",
]

# Full Activity form — every `data-id="control-id_<X>"` found on the
# Activities & Referrals Add-New dialog. Use as a reference when building
# higher-level helpers.
ACTIVITY_FIELDS = [
    "Subject",                  # textarea — free text
    "ActionUuid",               # dropdown — activity action
    "AgencyUuid",               # dropdown
    "ProviderUuid",             # dropdown
    "SubproviderUuid",          # dropdown
    "LevelCareLocusCareUuid",   # dropdown — Care Program
    "SiteUuid",                 # dropdown
    "StatusCodeUuid",           # dropdown — Status
    "ReasonCodeUuid",           # dropdown
    "StatusDate",               # date
    "DueDate",                  # date
    "StartDate", "StartTime",
    "EndDate", "EndTime",
    "FollowupStatus",           # dropdown
    "FollowupDate", "FollowupTime",
]

# Journal Add-New form (harvested 2026-04-24). The "Comments" RichTextBox
# has no data-id — fall back to fill_by_label("Comments", ...).
JOURNAL_FIELDS = [
    "Subject",                  # textarea
    "JournalTypeUuid",          # dropdown — defaults to "Progress Notes"
    "EntryDate",                # date textarea (placeholder "Enter date")
    "EntryTime",                # time textarea (placeholder "Enter time")
]

# Call Add-New form (16 controls).
CALL_FIELDS = [
    "StartDate",
    "TempEndDate", "TempSecondsPaused",  # timer-managed by UI usually
    "IsComplete",
    "CallTypeUuid",
    "CallerTypeUuid",
    "CallerConsumerUuid",
    "ConsumerUuid",
    "ReferredByTypeUuid",
    "PriorityTypeUuid",
    "PrimaryPaymentSourceTypeUuid",
    "AgeGroupUuid",
    "ConsumerGender",
    "ConsumerGenderIdentity",
    "DisabilitiesBrowseDisplay",
    "CallTopicCount",
]

# Care Plan Add-New form (10 controls).
CARE_PLAN_FIELDS = [
    "AgencyUuid",
    "LocusCareHistoryUuid",        # Care Program history link
    "PrimaryCareManagerProviderRoleTypeUuid",
    "PrimaryCareManagerProviderUuid",
    "StartDate", "EndDate",
    "StatusCodeUuid", "StatusDate",
    "ReasonCodeUuid",
    "PriorAuthorizationID",
]

# File Attachment Add-New form (5 controls).
FILE_ATTACHMENT_FIELDS = [
    "FolderUuid",
    "Description",
    "FileType", "FileSize",
    "DocumentBlob",                # the actual uploaded file
]

# Service Delivery Add-New form (19 controls).
SERVICE_DELIVERY_FIELDS = [
    "AgencyUuid", "ProviderUuid", "SubproviderUuid", "SiteUuid",
    "LocusCareHistoryUuid",
    "ServiceCategoryUuid", "ServiceUuid", "SubserviceUuid",
    "ServiceStartDateDisplay",
    "UnitType", "UnitPrice", "Units", "TotalCost",
    "FundIdentifierUuid",
    "DiagnosisCode",
    "PlaceOfServiceCodeBindingHelper",
    "ClientConsumerProviderUuid", "RecipientConsumerProviderUuid",
    "ServiceOrderNumber",
]

# Service Order Add-New form (10 controls).
SERVICE_ORDER_FIELDS = [
    "ServiceOrderID",
    "AgencyUuid", "ProviderUuid", "SubproviderUuid",
    "LocusCareHistoryUuid",
    "EffectiveDateDisplay", "ExpirationDate",
    "TotalOrderUnits", "TotalCost",
    "OrderItem",
]

# Care Enrollment dialog — opens by clicking the "Add New" button INSIDE
# the "Care Enrollments" section on the Details tab (NOT the top-toolbar
# Add New). Use `add_care_enrollment(...)` which finds it correctly.
CARE_ENROLLMENT_FIELDS = [
    "LevelCareLocusCareUuid",   # Care Program
    "LevelCareUuid",            # Level of Care
    "LocusCareUuid",            # Locus of Care
    "StatusCodeUuid",           # Status (dropdown)
    "ReasonCodeUuid",           # Reason (dropdown)
    "StatusDate",
    "ApplicationDate", "ReceivedDate",
    "StartDate", "EndDate", "TerminationDate",
]

# Assessment "New Assessment" form-picker dialog (15 controls).
# Note: this is the METADATA dialog. Clicking OK then opens the actual
# assessment form (which has its own hundreds of question fields, not
# yet mapped). Use add_assessment() for the metadata pass; capture the
# inner form separately if you need to drive the question answers.
ASSESSMENT_FIELDS = [
    "AssessFormFileName",          # dropdown — pick the assessment form
    "ShowFormsInSubfolders",       # checkbox
    "LevelCareLocusCareUuid",      # Care Program (defaults sensibly)
    "AgencyUuid",                  # defaults to CBES
    "ProviderUuid", "SubproviderUuid", "SiteUuid",
    "SessionDate",                 # date — assessment date
    "NextSessionDate",             # date — next assessment due
    "AssessorName",                # textarea
    "Author", "LastUpdated", "Version",  # mostly auto-populated
    "PasswordDisplay", "VerifyPassword", # for restricted forms
    # Comments has no data-id — use fill_by_label("Comments", ...)
]

# Details/Eligibility tab — nested-grid `data-id` prefixes. Each *Row is the
# row template; the others are columns you can read/write per row.
DETAILS_GRIDS = {
    "careProgram":   ["list_careProgramRow", "list_careProgram", "list_careProgramStatus",
                       "list_careProgramStartDate", "list_careProgramEndDate",
                       "list_reasonCodeUuid", "list_createOrganization"],
    "ethnicRaces":   ["list_ethnicRacesRow", "list_ethnicRaceGroupCode", "list_ethnicRaceNationality"],
    "providers":     ["list_providerSummaryRow", "list_providerUuid",
                       "list_providerStartDate", "list_providerEndDate"],
}

# Row-level icon buttons (Open / Delete / Print on each grid row).
ROW_BUTTONS = {
    "Open":   "ButtonOpenStyleImage",
    "Delete": "ButtonDeleteStyleImage",
    "Print":  "ButtonPrintStyleImage",
}


# ────────────────────────────────────────────────────────────────────
# CONFIG DATA CLASSES
# ────────────────────────────────────────────────────────────────────

@dataclass
class WellSkyConfig:
    username: str = "CBES5"
    password: str = "REDACTED"
    okta_url: str = OKTA_URL
    aging_url: str = AGING_URL
    default_timeout: int = 40
    headless: bool = False
    chrome_options: list[str] = field(default_factory=list)


# ────────────────────────────────────────────────────────────────────
# CORE CLIENT
# ────────────────────────────────────────────────────────────────────

class WellSkyClient:
    """Thin wrapper around a Chrome/Selenium session that knows how to
    navigate WellSky's two-layer UI (Angular shell + OpenSilver iframe)."""

    def __init__(self, config: Optional[WellSkyConfig] = None, **kwargs):
        self.config = config or WellSkyConfig(**kwargs)
        options = Options()
        if self.config.headless:
            options.add_argument("--headless=new")
        for arg in self.config.chrome_options:
            options.add_argument(arg)
        self.driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=options,
        )
        self.wait = WebDriverWait(self.driver, self.config.default_timeout)
        self._in_iframe = False

    # ── lifecycle ───────────────────────────────────────────────

    def close(self):
        try:
            self.driver.quit()
        except Exception:
            pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()

    # ── login ───────────────────────────────────────────────────

    def login(self, dashboard_timeout: int = 120):
        """Full Okta → ADFS → SAMS handoff. Returns when the dashboard
        iframe is ready (global Search box is clickable).

        The final wait is bumped to 120s by default because the SAMS
        iframe can take a long time to bootstrap (especially on the first
        login of the day). Pass `dashboard_timeout=` to override."""
        d = self.driver
        d.get(self.config.okta_url)

        username = self.wait.until(
            EC.presence_of_element_located((By.NAME, "identifier"))
        )
        username.clear()
        username.send_keys(self.config.username)
        d.find_element(By.CSS_SELECTOR, "input[type='submit']").click()

        adfs_next = self.wait.until(
            EC.element_to_be_clickable((By.ID, "nextButton"))
        )
        d.execute_script("arguments[0].click();", adfs_next)

        password = self.wait.until(
            EC.presence_of_element_located((By.ID, "passwordInput"))
        )
        password.clear()
        password.send_keys(self.config.password)
        d.find_element(By.ID, "submitButton").click()

        time.sleep(5)
        # Okta sometimes lands on its "My Apps" dashboard instead of deep-
        # linking into SAMS, and a first GET of the aging URL can hang past the
        # driver's HTTP timeout (which then kills chromedriver). So: bounded
        # page-load timeout, and up to three attempts to reach the dashboard.
        try:
            d.set_page_load_timeout(90)
        except Exception:
            pass
        last_err = None
        for attempt in range(3):
            try:
                d.get(self.config.aging_url)
            except TimeoutException as e:      # page load past 90s -> retry
                last_err = e
            try:
                WebDriverWait(d, 60 if attempt < 2 else dashboard_timeout).until(
                    EC.element_to_be_clickable(SEL_GLOBAL_SEARCH)
                )
                return
            except TimeoutException as e:
                last_err = e
                time.sleep(5)
        try:
            raise last_err
        except TimeoutException:
            # Dump state so we can debug — screenshot + url + title
            try:
                d.save_screenshot("login_failure.png")
            except Exception:
                pass
            raise TimeoutException(
                f"Dashboard didn't load within {dashboard_timeout}s.\n"
                f"  URL  : {d.current_url}\n"
                f"  Title: {d.title}\n"
                f"  Screenshot: login_failure.png (if saved)\n"
                "Common causes: Okta MFA push pending, ADFS re-auth required, "
                "or SAMS is cold-booting. Try pressing the push, then re-run."
            )

    # ── outer-shell navigation ─────────────────────────────────

    def to_default_content(self):
        """Leave any iframe we're in and return to the outer Angular shell."""
        self.driver.switch_to.default_content()
        self._in_iframe = False

    def goto_home(self):
        """Return to the Aging SPA landing page."""
        self.to_default_content()
        self.driver.get(self.config.aging_url)
        self.wait.until(EC.element_to_be_clickable(SEL_GLOBAL_SEARCH))

    def topnav(self, label: str):
        """Click a top-nav button by visible label. Opens the overflow
        menu first if needed (items like Calls/Reports/Administrator
        collapse there on narrower screens)."""
        self.to_default_content()
        xpath = TOPNAV_BUTTONS.get(label)
        if xpath is None:
            raise ValueError(f"Unknown top-nav label: {label}")
        try:
            btn = self.driver.find_element(By.XPATH, xpath)
        except NoSuchElementException:
            # try overflow
            self._open_overflow()
            btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        btn.click()

    def _open_overflow(self):
        """Open the top-nav kebab menu."""
        overflow = self.driver.find_element(
            By.XPATH, "//button[.//img[normalize-space()='more_vert']]"
        )
        overflow.click()

    def open_account_menu(self):
        """Open the account_circle dropdown in the top-right."""
        self.to_default_content()
        self.driver.find_element(
            By.XPATH, "//button[.//img[normalize-space()='account_circle']]"
        ).click()

    def sign_out(self):
        self.open_account_menu()
        self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@role='menuitem'][normalize-space()='Sign out']")
        )).click()

    # ── consumer search / open ─────────────────────────────────

    def open_consumer(self, consumer_id: str, header_timeout: int = 90):
        """Search by consumer ID, click the matching result, and switch
        into the OpenSilver iframe so subsequent helpers target the
        consumer record. The final wait for the record header to appear
        defaults to 90s — SAMS can be slow to bootstrap on a fresh
        session."""
        self.goto_home()
        search = self.wait.until(EC.element_to_be_clickable(SEL_GLOBAL_SEARCH))
        time.sleep(1.5)  # SAMS search autocomplete debounces hard
        search.click()
        search.clear()
        search.send_keys(str(consumer_id))

        option = self.wait.until(EC.presence_of_element_located(
            (By.XPATH, f"//mat-option[.//div[contains(text(),'ID: {consumer_id}')]]")
        ))
        ActionChains(self.driver).move_to_element(option).perform()
        option.click()

        iframe = self.wait.until(EC.presence_of_element_located(SEL_OUTER_IFRAME))
        self.driver.switch_to.frame(iframe)
        self._in_iframe = True

        # confirm the record header loaded — patient because SAMS is slow
        WebDriverWait(self.driver, header_timeout).until(
            EC.presence_of_element_located(
                (By.XPATH, f"//*[contains(text(),'{consumer_id}')]")
            )
        )

    # ── iframe-side helpers ────────────────────────────────────
    # Every helper below assumes the driver is already inside the SAMS
    # iframe. The `open_consumer` call puts you there. If you've bounced
    # out, call `enter_iframe()` again.

    def enter_iframe(self):
        self.to_default_content()
        iframe = self.wait.until(EC.presence_of_element_located(SEL_OUTER_IFRAME))
        self.driver.switch_to.frame(iframe)
        self._in_iframe = True

    # ── sidebar tab navigation ─────────────────────────────────

    def goto_tab(self, *text_parts: str, wait_after: float = 3.0):
        """Click a consumer sidebar tab. Pass one or more substrings — all
        must appear in the rendered span text. e.g.:
            goto_tab('Activities', 'Referrals')   # Activities & Referrals
            goto_tab('Details')
            goto_tab('Care', 'Plans')

        OpenSilver draws the tabs as text spans layered under an SVG
        overlay rect; a normal Selenium click gets intercepted. We try
        the normal click first and fall back to a JS click so the event
        still reaches the handler bound to the span."""
        conditions = " and ".join(
            f"contains(text(), '{t}')" for t in text_parts
        )
        xpath = f"//span[{conditions}]"
        self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        # SAMS keeps every consumer you opened this session as its own window,
        # stacked in the same DOM. The FIRST match is the oldest window's tab
        # (hidden -> "element not interactable"); the newest window is appended
        # last, so prefer the LAST displayed match. (2026-09-10: this broke
        # every 2nd+ consumer of a production run.)
        cands = self.driver.find_elements(By.XPATH, xpath)
        shown = [el for el in cands if self._is_shown(el)]
        tab = (shown or cands)[-1]
        self._click_resilient(tab)
        time.sleep(wait_after)

    @staticmethod
    def _is_shown(el) -> bool:
        try:
            if not el.is_displayed():
                return False
            r = el.rect
            return bool(r and (r.get("width", 0) or r.get("height", 0)))
        except Exception:
            return False

    def _click_resilient(self, el: WebElement):
        """Click an element; if the OpenSilver SVG overlay intercepts,
        retry via JS. Scrolls into view first."""
        try:
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center',inline:'center'});", el
            )
        except Exception:
            pass
        try:
            ActionChains(self.driver).move_to_element(el).perform()
        except Exception:
            pass
        try:
            el.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", el)

    # ── generic action buttons (Add New, Save, Delete, etc.) ──

    def click_button(self, label: str, strict: bool = True):
        """Click any button/span with the given visible label. If strict,
        only an exact match counts; otherwise substring.

        OpenSilver sometimes reports `is_displayed()=False` for elements
        that are in fact visible (the overlay makes the bounding rect
        tiny). We prefer displayed matches but fall back to any match
        with a non-zero rect — JS-click will still reach the handler."""
        if strict:
            xpath = f"//*[normalize-space()='{label}']"
        else:
            xpath = f"//*[contains(normalize-space(),'{label}')]"
        candidates = self.driver.find_elements(By.XPATH, xpath)
        # Newest consumer window last in the DOM -> try candidates last-first so
        # a stacked older window's identical button never wins (2026-09-10).
        candidates = list(reversed(candidates))
        # First pass: prefer Selenium-visible candidates
        for el in candidates:
            try:
                if el.is_displayed():
                    self._click_resilient(el)
                    return el
            except Exception:
                continue
        # Second pass: any candidate with a non-zero bounding rect
        for el in candidates:
            try:
                rect = el.rect
                if rect and (rect.get("width", 0) or rect.get("height", 0)):
                    self._click_resilient(el)
                    return el
            except Exception:
                continue
        raise NoSuchElementException(f"No visible element matching label: {label}")

    def click_add_new(self):
        return self.click_button("Add New")

    def click_section_add(self, section_label: str):
        """Click the 'Add New' button BELONGING to a labeled section
        (e.g. 'Care Enrollments' on the Details tab). The Details page
        renders multiple Add New icons — one per nested grid — and the
        top-toolbar Add New is yet another. This helper finds the Add
        New whose y-position is closest below the named section header,
        walks up to its `cursor: pointer` wrapper, and fires a real
        mousedown→mouseup→click sequence (the simple `.click()` doesn't
        trigger OpenSilver's button handlers for these section buttons)."""
        script = r"""
        const headerText = arguments[0];
        const header = [...document.querySelectorAll('span, div')]
          .find(el => (el.innerText || '').trim() === headerText && el.offsetParent);
        if (!header) return { error: 'header_not_found' };
        const hY = header.getBoundingClientRect().y;

        let best = null, bestDist = Infinity;
        document.querySelectorAll('span, div').forEach(el => {
          if (el.offsetParent === null) return;
          if ((el.innerText || '').trim() !== 'Add New') return;
          const r = el.getBoundingClientRect();
          if (r.y < hY - 5) return;
          if (r.y - hY < bestDist) { bestDist = r.y - hY; best = el; }
        });
        if (!best) return { error: 'no_add_below' };

        let node = best;
        for (let i = 0; i < 8 && node; i++) {
          const cs = getComputedStyle(node);
          if (cs.cursor === 'pointer' || node.getAttribute('role') === 'button') {
            const r = node.getBoundingClientRect();
            const cx = r.x + r.width / 2, cy = r.y + r.height / 2;
            ['mousedown', 'mouseup', 'click'].forEach(ev => {
              node.dispatchEvent(new MouseEvent(ev, {
                bubbles: true, cancelable: true,
                clientX: cx, clientY: cy,
                view: window, button: 0,
              }));
            });
            return { ok: true, depth: i, dist: bestDist };
          }
          node = node.parentElement;
        }
        return { error: 'no_clickable_parent' };
        """
        result = self.driver.execute_script(script, section_label)
        if result.get("error"):
            raise NoSuchElementException(
                f"Could not click Add New for section {section_label!r}: "
                f"{result['error']}"
            )
        time.sleep(2)
        return result

    def click_save(self):
        return self.click_button("Save")

    def click_save_and_close(self):
        return self.click_button("Save and Close")

    def commit_dialog(self):
        """Best-effort 'commit this dialog' — section-add dialogs use 'OK',
        record-level forms use 'Save and Close', some use plain 'Save'.
        We try them in order and return whichever worked."""
        for label in ("Save and Close", "OK", "Save"):
            try:
                self.click_button(label)
                return label
            except NoSuchElementException:
                continue
        raise NoSuchElementException(
            "No commit button found (tried 'Save and Close', 'OK', 'Save')"
        )

    def click_cancel(self):
        return self.click_button("Cancel")

    def click_delete(self):
        return self.click_button("Delete")

    def click_edit(self):
        return self.click_button("Edit")

    # ── form field helpers (keyed by stable data-id suffix) ───
    # Pass the property name without the `control-id_` prefix — e.g.
    # for <div data-id="control-id_Subject">, call:  fill_textarea("Subject", "...")

    @staticmethod
    def _control_xpath(property_name: str) -> str:
        return f"//div[@data-id='control-id_{property_name}']"

    def _find_control(self, property_name: str) -> WebElement:
        return self.wait.until(EC.presence_of_element_located(
            (By.XPATH, self._control_xpath(property_name))
        ))

    def fill_textarea(self, property_name: str, value: str, clear_first: bool = True):
        """Set a text/textarea control. WellSky wraps most single-line and
        multi-line inputs as <textarea> inside a data-id'd div. OpenSilver
        textareas are often zero-size / covered by an SVG overlay, so
        `el.send_keys()` fails with 'not interactable'. We focus via JS
        and then type via ActionChains (keyboard-level, bypasses the
        interactability check)."""
        xpath = self._control_xpath(property_name) + "//textarea"
        el = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        # Focus via JS — works even when the textarea is invisible
        self.driver.execute_script("arguments[0].focus();", el)
        time.sleep(0.3)
        if clear_first:
            # Select-all + delete via keyboard
            ActionChains(self.driver).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).send_keys(Keys.DELETE).perform()
            time.sleep(0.2)
        ActionChains(self.driver).send_keys(value).perform()

    def set_date(self, property_name: str, mmddyyyy: str):
        """Dates in WellSky are entered as textareas — same mechanic as
        fill_textarea, exposed as a named helper for readability."""
        self.fill_textarea(property_name, mmddyyyy)

    def pick_dropdown(self, property_name: str, value: str,
                      press_enter: bool = True, settle: float = 1.0):
        """Open an OpenSilver dropdown (they're not native <select>), type
        to filter, optionally press ENTER to accept."""
        anchor = self._find_control(property_name)
        self._click_resilient(anchor)
        time.sleep(0.5)
        ActionChains(self.driver).send_keys(value).perform()
        time.sleep(settle)
        if press_enter:
            ActionChains(self.driver).send_keys(Keys.ENTER).perform()
            time.sleep(settle)

    def set_checkbox(self, property_name: str, checked: bool = True):
        """Toggle a check-box control. OpenSilver checkboxes are
        click-driven DIVs (often without aria-checked, so we look for
        common visual indicators). Falls back to JS-click on intercept."""
        anchor = self._find_control(property_name)
        current = anchor.get_attribute("aria-checked")
        if current is None:
            # Some OpenSilver checkboxes use a child marker — toggling is
            # generally idempotent-friendly when we just click once.
            self._click_resilient(anchor)
            return
        if (current == "true") != checked:
            self._click_resilient(anchor)

    def select_radio(self, property_name: str, option_label: str):
        """Click a radio option. The data-id anchors the radio group; the
        option is a descendant span with the label text."""
        xpath = (
            self._control_xpath(property_name)
            + f"//*[normalize-space()='{option_label}']"
        )
        self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()

    def read_field(self, property_name: str) -> str:
        """Read back the current text inside a data-id control (useful for
        assertions in tests)."""
        el = self._find_control(property_name)
        return (el.text or "").strip()

    def list_dropdown_options(self, property_name: str,
                               scroll_passes: int = 8,
                               close_after: bool = True) -> list[str]:
        """Open an OpenSilver dropdown, scroll through its popup, and
        return every option visible along the way. Closes the popup via
        Escape unless `close_after=False`.

        The popup virtualises (only renders a viewport-worth of items),
        so we scroll with PageDown/End and diff the DOM on each pass.
        `scroll_passes` caps how many scrolls we do — bump it for huge
        lists (Providers has hundreds)."""
        anchor = self._find_control(property_name)
        before = set(self.driver.execute_script(
            "return [...document.querySelectorAll('span, div')]"
            ".filter(e=>e.offsetParent)"
            ".map(e=>(e.innerText||'').trim())"
            ".filter(t=>t && t.length<80 && t.length>1);"
        ))
        self._click_resilient(anchor)
        time.sleep(1.5)

        collected: set[str] = set()

        def snapshot() -> set[str]:
            return set(self.driver.execute_script(
                "return [...document.querySelectorAll('span, div')]"
                ".filter(e=>e.offsetParent)"
                ".map(e=>(e.innerText||'').trim())"
                ".filter(t=>t && t.length<80 && t.length>1);"
            ))

        # Initial visible items
        collected |= snapshot() - before

        # Scroll through the popup
        for _ in range(scroll_passes):
            ActionChains(self.driver).send_keys(Keys.PAGE_DOWN).perform()
            time.sleep(0.4)
            new = snapshot() - before
            if new.issubset(collected):
                # No progress — we've reached the end
                break
            collected |= new

        if close_after:
            ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.6)

        # Sort & drop newline-merged artifacts
        out = []
        for item in collected:
            for piece in item.splitlines():
                p = piece.strip()
                if p and len(p) < 80:
                    out.append(p)
        return sorted(set(out))

    def fill_by_label(self, label: str, value: str):
        """Escape hatch for fields without a `data-id` (e.g. Journal's
        Comments RichTextBox). Finds the <span> with the given label,
        walks up to the nearest container with a sibling textarea/input,
        and types into it.

        This is fragile by design — prefer `fill_textarea(...)` when a
        data-id exists. Use this for Comments, Note bodies, and similar
        rich-text fields where no stable hook exists.

        IMPORTANT: named fields (Subject, Entry Date, …) each live inside a
        `[data-id^="control-id_"]` wrapper and Subject is a <textarea>. A naive
        ancestor walk from the "Comments" label reaches the shared form
        container and grabs Subject's textarea first — dumping the whole note
        into Subject. So we EXCLUDE any editable contained in a control-id
        wrapper and take the label-less one (that's the Comments RichTextBox),
        preferring a contenteditable over a textarea."""
        script = r"""
        const label = arguments[0];
        const lbl = [...document.querySelectorAll('span')]
            .find(s => (s.innerText || '').trim() === label);
        if (!lbl) return null;
        const named = (el) => el.closest('[data-id^="control-id_"]') !== null;
        const pick = (root) => {
            if (!root || !root.querySelectorAll) return null;
            // Prefer the rich-text editor, then bare textareas/inputs.
            const order = ['[contenteditable="true"]', 'textarea', 'input[type="text"]'];
            for (const sel of order) {
                for (const ed of root.querySelectorAll(sel)) {
                    if (named(ed)) continue;               // skip Subject/EntryDate/etc.
                    if (ed.offsetParent === null &&
                        ed.getAttribute('contenteditable') !== 'true') continue;
                    return ed;
                }
            }
            return null;
        };
        let node = lbl;
        for (let i = 0; i < 10 && node; i++) {
            const ed = pick(node);
            if (ed) {
                ed.focus();
                return ed.getAttribute('contenteditable') === 'true'
                    ? 'CE' : ed.tagName;
            }
            node = node.parentElement;
        }
        return null;
        """
        tag = self.driver.execute_script(script, label)
        if tag is None:
            raise NoSuchElementException(f"Could not find a field near label: {label!r}")
        # After focus, clear any stale content then type into the active element.
        ae = self.driver.switch_to.active_element
        ActionChains(self.driver).key_down(Keys.CONTROL).send_keys("a") \
            .key_up(Keys.CONTROL).send_keys(Keys.DELETE).perform()
        time.sleep(0.2)
        ae.send_keys(value)

    # ── dialogs / confirmations ────────────────────────────────

    def confirm_yes(self):
        self.click_button("Yes")

    def confirm_ok(self):
        self.click_button("OK")

    def confirm_no(self):
        self.click_button("No")

    def dismiss_modal(self):
        # fallback: click an X or press Escape
        try:
            self.click_button("Close")
        except NoSuchElementException:
            self.driver.switch_to.active_element.send_keys(Keys.ESCAPE)

    # ── higher-level composed actions ──────────────────────────

    def add_activity(self,
                     subject: str,
                     action: str,
                     agency: Optional[str] = DEFAULT_AGENCY,
                     provider: Optional[str] = None,
                     subprovider: Optional[str] = None,
                     care_program: Optional[str] = None,
                     site: Optional[str] = None,
                     status: Optional[str] = None,
                     reason: Optional[str] = None,
                     status_date: Optional[str] = None,
                     due_date: Optional[str] = None,
                     start_date: Optional[str] = None,
                     start_time: Optional[str] = None,
                     end_date: Optional[str] = None,
                     end_time: Optional[str] = None,
                     followup_status: Optional[str] = None,
                     followup_date: Optional[str] = None,
                     followup_time: Optional[str] = None,
                     save: bool = False):
        """Create one Activity on the currently-open consumer. Assumes
        you're already on Activities & Referrals. Every field maps 1:1
        to a `control-id_<X>` on the Add-New dialog.

        The full field list was harvested from the sandbox (see
        ACTIVITY_FIELDS) — pass only the ones you care about."""
        self.click_add_new()
        self.fill_textarea("Subject", subject)
        self.pick_dropdown("ActionUuid", action)
        if agency:          self.pick_dropdown("AgencyUuid", agency)
        if provider:        self.pick_dropdown("ProviderUuid", provider)
        if subprovider:     self.pick_dropdown("SubproviderUuid", subprovider)
        if care_program:    self.pick_dropdown("LevelCareLocusCareUuid", care_program)
        if site:            self.pick_dropdown("SiteUuid", site)
        if status:          self.pick_dropdown("StatusCodeUuid", status)
        if reason:          self.pick_dropdown("ReasonCodeUuid", reason)
        if status_date:     self.set_date("StatusDate", status_date)
        if due_date:        self.set_date("DueDate", due_date)
        if start_date:      self.set_date("StartDate", start_date)
        if start_time:      self.fill_textarea("StartTime", start_time)
        if end_date:        self.set_date("EndDate", end_date)
        if end_time:        self.fill_textarea("EndTime", end_time)
        if followup_status: self.pick_dropdown("FollowupStatus", followup_status)
        if followup_date:   self.set_date("FollowupDate", followup_date)
        if followup_time:   self.fill_textarea("FollowupTime", followup_time)
        if save:
            self.click_save_and_close()

    # ── cleanup ────────────────────────────────────────────────

    def close_any_dialog(self):
        """Best-effort: close an open Add/Edit dialog. Only tries 'Cancel'
        (the dialog-scoped button). DOES NOT try 'Close' — that's the
        record-level button and would close the whole consumer."""
        try:
            self.click_button("Cancel")
            time.sleep(1.5)
        except NoSuchElementException:
            return
        # Some Cancels trigger a "Discard changes?" modal
        for confirm in ("Yes", "OK", "Discard"):
            try:
                self.click_button(confirm)
                time.sleep(1)
                return
            except NoSuchElementException:
                continue

    def add_journal(self,
                    subject: str,
                    journal_type: Optional[str] = None,
                    entry_date: Optional[str] = None,
                    entry_time: Optional[str] = None,
                    comments: Optional[str] = None,
                    save: bool = False):
        """Create a Journal entry. Assumes you're on the Journals tab.

        Fields (all optional except subject):
            - journal_type : dropdown, defaults in UI to 'Progress Notes'
            - entry_date   : 'MM/DD/YYYY'
            - entry_time   : 'HH:MM AM/PM'
            - comments     : free-text rich-text body (no data-id; uses
                             fill_by_label)."""
        self.click_add_new()
        self.fill_textarea("Subject", subject)
        if journal_type:
            self.pick_dropdown("JournalTypeUuid", journal_type)
        if entry_date:
            self.set_date("EntryDate", entry_date)
        if entry_time:
            self.fill_textarea("EntryTime", entry_time)
        if comments:
            self.fill_by_label("Comments", comments)
        if save:
            self.click_save_and_close()

    def add_care_enrollment(self,
                            care_program: str,
                            start_date: str,
                            status: Optional[str] = None,
                            level_of_care: Optional[str] = None,
                            locus_of_care: Optional[str] = None,
                            status_date: Optional[str] = None,
                            application_date: Optional[str] = None,
                            received_date: Optional[str] = None,
                            end_date: Optional[str] = None,
                            termination_date: Optional[str] = None,
                            reason: Optional[str] = None,
                            save: bool = False):
        """Create a Care Enrollment row in the Details → Care Enrollments
        section. You must already be on the Details tab (the default
        landing for an open consumer).

        Note: the Add New button for Care Enrollments is the SECTION-level
        one, not the top-toolbar — this helper locates the right one by
        walking down from the 'Care Enrollments' header span and firing
        a synthetic mouse-event sequence (which OpenSilver requires)."""
        self.click_section_add("Care Enrollments")
        time.sleep(2)
        self.pick_dropdown("LevelCareLocusCareUuid", care_program)
        if level_of_care:    self.pick_dropdown("LevelCareUuid", level_of_care)
        if locus_of_care:    self.pick_dropdown("LocusCareUuid", locus_of_care)
        if status:           self.pick_dropdown("StatusCodeUuid", status)
        if reason:           self.pick_dropdown("ReasonCodeUuid", reason)
        if status_date:      self.set_date("StatusDate", status_date)
        if application_date: self.set_date("ApplicationDate", application_date)
        if received_date:    self.set_date("ReceivedDate", received_date)
        self.set_date("StartDate", start_date)
        if end_date:         self.set_date("EndDate", end_date)
        if termination_date: self.set_date("TerminationDate", termination_date)
        if save:
            # Section-add dialogs typically use OK (not Save and Close).
            # commit_dialog tries the right ones in order.
            self.commit_dialog()

    def open_care_enrollment_row(self, care_program: str,
                                  active_only: bool = True) -> bool:
        """Open the edit dialog for the Care Enrollments row whose Care
        Program matches `care_program`. Pass `active_only=True` (default)
        to only match rows with an empty End Date.

        Returns True if a row was opened. Raises NoSuchElementException
        if no matching row exists.

        Identifies the row by inspecting the visible grid (the
        underlying `care_enrollment_uuid` is internal-only and never
        renders in the UI)."""
        script = r"""
        const target = arguments[0];
        const activeOnly = arguments[1];
        const rows = document.querySelectorAll('[data-id="list_careProgramRow"]');
        for (const row of rows) {
          if (row.offsetParent === null) continue;
          const progCell = row.querySelector('[data-id="list_careProgram"]');
          const endCell = row.querySelector('[data-id="list_careProgramEndDate"]');
          const prog = (progCell?.innerText || '').trim();
          const end = (endCell?.innerText || '').trim();
          if (prog !== target) continue;
          // WellSky renders 'no end date' as '(Not Specified)' (sometimes -).
          const ACTIVE_SENTINELS = ['', ' ', '(Not Specified)', 'Not Specified', '-'];
          if (activeOnly && !ACTIVE_SENTINELS.includes(end)) continue;
          // Open icons are siblings (not children) of the row — match by Y.
          const rowRect = row.getBoundingClientRect();
          const rowY = rowRect.y;
          const opens = document.querySelectorAll('[data-id="ButtonOpenStyleImage"]');
          let openBtn = null;
          for (const b of opens) {
            if (b.offsetParent === null) continue;
            const r = b.getBoundingClientRect();
            if (Math.abs(r.y - rowY) < 5) { openBtn = b; break; }
          }
          if (!openBtn) return { error: 'no_open_button_at_row_Y_' + Math.round(rowY) };
          // Walk up to the cursor:pointer parent
          let node = openBtn;
          for (let i = 0; i < 8 && node; i++) {
            const cs = getComputedStyle(node);
            if (cs.cursor === 'pointer' || node.getAttribute('role') === 'button') {
              node.scrollIntoView({block: 'center'});
              const r = node.getBoundingClientRect();
              const cx = r.x + r.width/2, cy = r.y + r.height/2;
              ['mousedown','mouseup','click'].forEach(ev => {
                node.dispatchEvent(new MouseEvent(ev, {
                  bubbles: true, cancelable: true,
                  clientX: cx, clientY: cy, view: window, button: 0,
                }));
              });
              return { ok: true, program: prog, end_date: end };
            }
            node = node.parentElement;
          }
          // fall back to synth-click on the icon itself
          const r = openBtn.getBoundingClientRect();
          ['mousedown','mouseup','click'].forEach(ev => {
            openBtn.dispatchEvent(new MouseEvent(ev, {
              bubbles: true, cancelable: true,
              clientX: r.x + r.width/2, clientY: r.y + r.height/2,
              view: window, button: 0,
            }));
          });
          return { ok: true, program: prog, end_date: end, fallback: true };
        }
        return { error: 'no_matching_row' };
        """
        result = self.driver.execute_script(script, care_program, active_only)
        if result.get("error"):
            raise NoSuchElementException(
                f"Could not open Care Enrollment row for {care_program!r}: "
                f"{result['error']}"
            )
        time.sleep(3)
        return True

    def disenroll_care_program(self,
                                care_program: str,
                                end_date: str,
                                termination_date: Optional[str] = None,
                                status: Optional[str] = None,
                                reason: Optional[str] = None,
                                status_date: Optional[str] = None,
                                save: bool = False):
        """Disenroll a consumer from a Care Program. Finds the active
        enrollment row by program name, opens it, sets EndDate (and
        optionally TerminationDate / Status / Reason), and commits.

        Identifies the row by visible Care Program name + empty End Date.
        If the consumer has multiple historical enrollments in the same
        program, only the active (no-end-date) row is matched."""
        self.open_care_enrollment_row(care_program, active_only=True)
        time.sleep(2)
        self.set_date("EndDate", end_date)
        if termination_date: self.set_date("TerminationDate", termination_date)
        if status_date:      self.set_date("StatusDate", status_date)
        if status:           self.pick_dropdown("StatusCodeUuid", status)
        if reason:           self.pick_dropdown("ReasonCodeUuid", reason)
        if save:
            self.commit_dialog()

    def add_assessment(self,
                       form_name: str,
                       session_date: str,
                       assessor_name: Optional[str] = None,
                       care_program: Optional[str] = None,
                       agency: Optional[str] = DEFAULT_AGENCY,
                       provider: Optional[str] = None,
                       subprovider: Optional[str] = None,
                       site: Optional[str] = None,
                       next_session_date: Optional[str] = None,
                       comments: Optional[str] = None,
                       show_subfolders: bool = False,
                       password: Optional[str] = None,
                       open_form: bool = False):
        """Open the New Assessment metadata dialog. On the Assessments tab.

        `form_name` is the assessment template to use (e.g. 'NAPIS - Title III',
        'RAMWP', 'Initial Assessment'). Pick from the AssessFormFileName
        dropdown — call list_dropdown_options('AssessFormFileName') for the
        live list.

        Set `open_form=True` to click OK and open the assessment itself.
        Otherwise the metadata is filled but the dialog stays open so you
        can review before committing. Driving the actual assessment form
        (the question answers) is NOT handled here — that's a separate
        screen with its own fields."""
        self.click_add_new()
        if show_subfolders:
            self.set_checkbox("ShowFormsInSubfolders", True)
        self.pick_dropdown("AssessFormFileName", form_name)
        if care_program:
            self.pick_dropdown("LevelCareLocusCareUuid", care_program)
        if agency:
            self.pick_dropdown("AgencyUuid", agency)
        if provider:
            self.pick_dropdown("ProviderUuid", provider)
        if subprovider:
            self.pick_dropdown("SubproviderUuid", subprovider)
        if site:
            self.pick_dropdown("SiteUuid", site)
        self.set_date("SessionDate", session_date)
        if next_session_date:
            self.set_date("NextSessionDate", next_session_date)
        if assessor_name:
            self.fill_textarea("AssessorName", assessor_name)
        if password:
            self.fill_textarea("PasswordDisplay", password)
            self.fill_textarea("VerifyPassword", password)
        if comments:
            self.fill_by_label("Comments", comments)
        if open_form:
            self.click_button("OK")

    def add_call(self,
                 start_date: str,
                 call_type: Optional[str] = None,
                 caller_type: Optional[str] = None,
                 priority: Optional[str] = None,
                 referred_by: Optional[str] = None,
                 complete: Optional[bool] = None,
                 save: bool = False):
        """Create a Call record. On the Calls tab."""
        self.click_add_new()
        self.set_date("StartDate", start_date)
        if call_type:    self.pick_dropdown("CallTypeUuid", call_type)
        if caller_type:  self.pick_dropdown("CallerTypeUuid", caller_type)
        if priority:     self.pick_dropdown("PriorityTypeUuid", priority)
        if referred_by:  self.pick_dropdown("ReferredByTypeUuid", referred_by)
        if complete is not None:
            self.set_checkbox("IsComplete", complete)
        if save:
            self.click_save_and_close()

    def add_care_plan(self,
                      start_date: str,
                      end_date: Optional[str] = None,
                      agency: Optional[str] = DEFAULT_AGENCY,
                      care_program: Optional[str] = None,
                      primary_cm_role: Optional[str] = None,
                      primary_cm: Optional[str] = None,
                      status: Optional[str] = None,
                      status_date: Optional[str] = None,
                      reason: Optional[str] = None,
                      prior_auth_id: Optional[str] = None,
                      save: bool = False):
        """Create a Care Plan. On the Care Plans tab."""
        self.click_add_new()
        self.set_date("StartDate", start_date)
        if end_date:        self.set_date("EndDate", end_date)
        if agency:          self.pick_dropdown("AgencyUuid", agency)
        if care_program:    self.pick_dropdown("LocusCareHistoryUuid", care_program)
        if primary_cm_role: self.pick_dropdown("PrimaryCareManagerProviderRoleTypeUuid", primary_cm_role)
        if primary_cm:      self.pick_dropdown("PrimaryCareManagerProviderUuid", primary_cm)
        if status:          self.pick_dropdown("StatusCodeUuid", status)
        if status_date:     self.set_date("StatusDate", status_date)
        if reason:          self.pick_dropdown("ReasonCodeUuid", reason)
        if prior_auth_id:   self.fill_textarea("PriorAuthorizationID", prior_auth_id)
        if save:
            self.click_save_and_close()

    def add_service_order(self,
                          effective_date: str,
                          expiration_date: Optional[str] = None,
                          agency: Optional[str] = DEFAULT_AGENCY,
                          provider: Optional[str] = None,
                          subprovider: Optional[str] = None,
                          care_program: Optional[str] = None,
                          order_id: Optional[str] = None,
                          total_units: Optional[str] = None,
                          total_cost: Optional[str] = None,
                          save: bool = False):
        """Create a Service Order. On the Service Orders tab."""
        self.click_add_new()
        if order_id:        self.fill_textarea("ServiceOrderID", order_id)
        self.set_date("EffectiveDateDisplay", effective_date)
        if expiration_date: self.set_date("ExpirationDate", expiration_date)
        if agency:          self.pick_dropdown("AgencyUuid", agency)
        if provider:        self.pick_dropdown("ProviderUuid", provider)
        if subprovider:     self.pick_dropdown("SubproviderUuid", subprovider)
        if care_program:    self.pick_dropdown("LocusCareHistoryUuid", care_program)
        if total_units:     self.fill_textarea("TotalOrderUnits", total_units)
        if total_cost:      self.fill_textarea("TotalCost", total_cost)
        if save:
            self.click_save_and_close()

    def add_service_delivery(self,
                             service_start_date: str,
                             service: Optional[str] = None,
                             service_category: Optional[str] = None,
                             subservice: Optional[str] = None,
                             agency: Optional[str] = DEFAULT_AGENCY,
                             provider: Optional[str] = None,
                             subprovider: Optional[str] = None,
                             site: Optional[str] = None,
                             care_program: Optional[str] = None,
                             units: Optional[str] = None,
                             unit_price: Optional[str] = None,
                             unit_type: Optional[str] = None,
                             total_cost: Optional[str] = None,
                             fund: Optional[str] = None,
                             diagnosis_code: Optional[str] = None,
                             save: bool = False):
        """Create a Service Delivery. On the Service Deliveries tab.
        Most fields are dropdowns; Units / UnitPrice / TotalCost are text."""
        self.click_add_new()
        self.set_date("ServiceStartDateDisplay", service_start_date)
        if service_category: self.pick_dropdown("ServiceCategoryUuid", service_category)
        if service:          self.pick_dropdown("ServiceUuid", service)
        if subservice:       self.pick_dropdown("SubserviceUuid", subservice)
        if agency:           self.pick_dropdown("AgencyUuid", agency)
        if provider:         self.pick_dropdown("ProviderUuid", provider)
        if subprovider:      self.pick_dropdown("SubproviderUuid", subprovider)
        if site:             self.pick_dropdown("SiteUuid", site)
        if care_program:     self.pick_dropdown("LocusCareHistoryUuid", care_program)
        if unit_type:        self.pick_dropdown("UnitType", unit_type)
        if units:            self.fill_textarea("Units", units)
        if unit_price:       self.fill_textarea("UnitPrice", unit_price)
        if total_cost:       self.fill_textarea("TotalCost", total_cost)
        if fund:             self.pick_dropdown("FundIdentifierUuid", fund)
        if diagnosis_code:   self.fill_textarea("DiagnosisCode", diagnosis_code)
        if save:
            self.click_save_and_close()

    def add_file_attachment(self,
                            folder: str,
                            description: str,
                            file_path: str,
                            save: bool = False):
        """Upload a File Attachment. On the File Attachments tab.
        `file_path` is the absolute path to the local file; the DocumentBlob
        control is a hidden file-input that we can send_keys to directly."""
        self.click_add_new()
        self.pick_dropdown("FolderUuid", folder)
        self.fill_textarea("Description", description)
        # DocumentBlob is a file input — send the path.
        blob_xpath = self._control_xpath("DocumentBlob") + "//input[@type='file']"
        upload = self.wait.until(EC.presence_of_element_located((By.XPATH, blob_xpath)))
        upload.send_keys(file_path)
        if save:
            self.click_save_and_close()

    # ── low-level escape hatches ───────────────────────────────

    def find(self, by: str, selector: str) -> WebElement:
        return self.driver.find_element(by, selector)

    def find_all(self, by: str, selector: str) -> list[WebElement]:
        return self.driver.find_elements(by, selector)

    def js(self, script: str, *args):
        return self.driver.execute_script(script, *args)

    def screenshot(self, path: str):
        self.driver.save_screenshot(path)
