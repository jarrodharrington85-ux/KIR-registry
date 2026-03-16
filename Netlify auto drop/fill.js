/**
 * Kiribati Registry Auto-Fill Tool
 * Phase 1: Core Field Setter + BN-0 Simple Fields
 * 
 * Target: ki.ua.paradigmapps.com/corp/filing.aspx
 * Stack: Vue 2.5.16 inside AngularJS 1.6.8, jQuery 1.12.4
 * Field IDs: {fieldname}_{UUID} — selector: [id^="fieldname_"]
 * Vue access: element.parentElement.__vue__ (vue-input component)
 * 
 * Delivery: bookmarklet → loads this script → prompt() for JSON → fills form
 */

(function() {
  'use strict';

  const VERSION = '0.2.0';
  const TOOL_NAME = 'KI Registry Auto-Fill';

  // ============================================================
  // LOGGING
  // ============================================================

  const log = {
    _entries: [],
    _counts: { ok: 0, skip: 0, fail: 0 },

    info(msg)  { console.log(`[AutoFill] ${msg}`); },
    warn(msg)  { console.warn(`[AutoFill] ⚠ ${msg}`); },
    error(msg) { console.error(`[AutoFill] ✗ ${msg}`); },

    field(name, status, detail) {
      const icon = status === 'ok' ? '✓' : status === 'skip' ? '–' : '✗';
      const entry = { name, status, detail };
      this._entries.push(entry);
      this._counts[status] = (this._counts[status] || 0) + 1;
      console.log(`[AutoFill]   ${icon} ${name}: ${detail || status}`);
    },

    summary() {
      const c = this._counts;
      const total = c.ok + c.skip + c.fail;
      console.log(`\n[AutoFill] ══════════════════════════════════`);
      console.log(`[AutoFill] SUMMARY: ${c.ok} filled, ${c.skip} skipped, ${c.fail} failed (${total} total)`);
      if (c.fail > 0) {
        console.log(`[AutoFill] Failed fields:`);
        this._entries.filter(e => e.status === 'fail').forEach(e => {
          console.log(`[AutoFill]   ✗ ${e.name}: ${e.detail}`);
        });
      }
      console.log(`[AutoFill] ══════════════════════════════════\n`);
      return this._counts;
    },

    reset() {
      this._entries = [];
      this._counts = { ok: 0, skip: 0, fail: 0 };
    }
  };

  // ============================================================
  // FIELD SETTER — Core module
  // ============================================================
  // 
  // Strategy: Set the native DOM value, then trigger events that
  // Vue 2's v-model listens on. Vue 2 uses input events on text
  // fields and change events on selects/radios. We also use the
  // native value setter to bypass any framework getter/setter
  // on the .value property.

  const fieldSetter = {

    /**
     * Find an element by field name prefix.
     * Field IDs are {fieldname}_{UUID}. We use [id^="name_"] to match.
     * Optionally scope to a container element (for grid rows).
     */
    find(fieldName, container) {
      const scope = container || document;
      // Primary: attribute starts-with selector
      const el = scope.querySelector(`[id^="${fieldName}_"]`);
      if (el) return el;
      // Fallback: exact ID match (some fields like corpofficerdescriptionofrelationship have no UUID)
      return scope.querySelector(`#${fieldName}`);
    },

    /**
     * Set a text/email/tel input value with full Vue 2 reactivity.
     */
    setText(fieldName, value, container) {
      if (value === undefined || value === null || value === '') return 'skip';
      const el = this.find(fieldName, container);
      if (!el) return 'not_found';

      try {
        // Use native setter to bypass any framework property descriptor
        const nativeSetter = Object.getOwnPropertyDescriptor(
          HTMLInputElement.prototype, 'value'
        )?.set || Object.getOwnPropertyDescriptor(
          HTMLTextAreaElement.prototype, 'value'
        )?.set;

        if (nativeSetter) {
          nativeSetter.call(el, value);
        } else {
          el.value = value;
        }

        // Trigger events Vue 2 listens on
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));

        // Also try to update via Vue instance directly
        this._setVue(el, value);

        return 'ok';
      } catch (e) {
        return 'error: ' + e.message;
      }
    },

    /**
     * Set a <select> value.
     * `value` is the option value (string). For country/nationality selects
     * this is the numeric option value from the registry's dropdown.
     */
    setSelect(fieldName, value, container) {
      if (value === undefined || value === null || value === '') return 'skip';
      const el = this.find(fieldName, container);
      if (!el) return 'not_found';

      try {
        el.value = String(value);

        // Verify option actually exists
        if (el.value !== String(value)) {
          // Try matching by option text (case-insensitive)
          const match = Array.from(el.options).find(
            o => o.text.trim().toLowerCase() === String(value).toLowerCase()
          );
          if (match) {
            el.value = match.value;
          } else {
            return `option_not_found: "${value}"`;
          }
        }

        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('input', { bubbles: true }));
        this._setVue(el, el.value);

        return 'ok';
      } catch (e) {
        return 'error: ' + e.message;
      }
    },

    /**
     * Set a <select> by matching the option text (label) rather than value.
     * Used when the JSON has human-readable country/nationality names.
     */
    setSelectByText(fieldName, text, container) {
      if (!text) return 'skip';
      const el = this.find(fieldName, container);
      if (!el) return 'not_found';

      try {
        const normalised = text.trim().toLowerCase();
        const match = Array.from(el.options).find(
          o => o.text.trim().toLowerCase() === normalised
        );
        if (!match) {
          // Try partial match
          const partial = Array.from(el.options).find(
            o => o.text.trim().toLowerCase().includes(normalised) ||
                 normalised.includes(o.text.trim().toLowerCase())
          );
          if (partial) {
            el.value = partial.value;
          } else {
            return `text_not_found: "${text}"`;
          }
        } else {
          el.value = match.value;
        }

        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('input', { bubbles: true }));
        this._setVue(el, el.value);

        return 'ok';
      } catch (e) {
        return 'error: ' + e.message;
      }
    },

    /**
     * Set a radio button by value.
     * Radio groups share a name attribute: {fieldname}_{UUID}.
     * We find the group by [name^="fieldname_"], then click the matching value.
     */
    setRadio(fieldName, value, container) {
      if (value === undefined || value === null || value === '') return 'skip';
      const scope = container || document;

      try {
        // Find all radios whose name starts with the field name
        const radios = scope.querySelectorAll(`[name^="${fieldName}_"]`);
        if (radios.length === 0) {
          // Fallback: try exact name match
          const exact = scope.querySelectorAll(`[name="${fieldName}"]`);
          if (exact.length === 0) return 'not_found';
          return this._clickRadio(exact, value);
        }
        return this._clickRadio(radios, value);
      } catch (e) {
        return 'error: ' + e.message;
      }
    },

    _clickRadio(radios, value) {
      const strVal = String(value);
      for (const r of radios) {
        if (r.value === strVal) {
          r.checked = true;
          r.dispatchEvent(new Event('change', { bubbles: true }));
          r.dispatchEvent(new Event('click', { bubbles: true }));
          // Also try the click() method which can trigger Vue watchers
          r.click();
          return 'ok';
        }
      }
      return `value_not_found: "${value}" in [${Array.from(radios).map(r => r.value).join(',')}]`;
    },

    /**
     * Set a textarea value (e.g. beneficial owner relationship field).
     */
    setTextarea(fieldName, value, container) {
      if (!value) return 'skip';
      const el = this.find(fieldName, container);
      if (!el) return 'not_found';

      try {
        const nativeSetter = Object.getOwnPropertyDescriptor(
          HTMLTextAreaElement.prototype, 'value'
        )?.set;
        if (nativeSetter) {
          nativeSetter.call(el, value);
        } else {
          el.value = value;
        }
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));
        this._setVue(el, value);
        return 'ok';
      } catch (e) {
        return 'error: ' + e.message;
      }
    },

    /**
     * Try to update the Vue component's localValue directly.
     * Access path: element.parentElement.__vue__ → vue-input component.
     */
    _setVue(el, value) {
      try {
        // Walk up to find the Vue instance
        let node = el;
        for (let i = 0; i < 5; i++) {
          if (node.__vue__) {
            const vm = node.__vue__;
            // vue-input components typically have a `localValue` or `value` data prop
            if ('localValue' in vm) {
              vm.localValue = value;
            }
            if (typeof vm.$emit === 'function') {
              vm.$emit('input', value);
              vm.$emit('change', value);
            }
            return true;
          }
          node = node.parentElement;
          if (!node) break;
        }
      } catch (e) {
        // Vue access is best-effort; DOM events are the primary mechanism
      }
      return false;
    }
  };

  // ============================================================
  // TAB NAVIGATION
  // ============================================================

  function navigateToTab(tabId) {
    const tabLink = document.querySelector(`a[href="#${tabId}"]`);
    if (tabLink) {
      tabLink.click();
      log.info(`Navigated to tab: #${tabId}`);
      return true;
    }
    log.warn(`Tab not found: #${tabId}`);
    return false;
  }

  // ============================================================
  // FORM DETECTION
  // ============================================================

  function detectForm() {
    // Check corporiginalcorptypeid — "3" = business name
    const typeEl = fieldSetter.find('corporiginalcorptypeid');
    const corpType = typeEl ? typeEl.value : null;

    // Check filing type
    const filingEl = fieldSetter.find('corpfilingcorpfilingtypeid');
    const filingType = filingEl ? filingEl.value : null;

    // Check for business name (read-only)
    const nameEl = fieldSetter.find('corporiginalname');
    const entityName = nameEl ? nameEl.value : null;

    const regEl = fieldSetter.find('corporiginalcorpregistrationnumber');
    const regNumber = regEl ? regEl.value : null;

    // Determine form type
    let formType = 'unknown';
    if (corpType === '3') {
      // BN form — check if it's BN-0 (has pre-populated name) or BN-1 (blank)
      if (entityName && entityName.trim() !== '') {
        formType = 'BN-0'; // Re-registration (has existing business name)
      } else {
        formType = 'BN-1'; // New registration (blank)
      }
    }

    const result = { formType, corpType, filingType, entityName, regNumber };
    log.info(`Detected form: ${formType} | Entity: ${entityName || '(none)'} | Reg#: ${regNumber || '(none)'}`);
    return result;
  }

  // ============================================================
  // COUNTRY/NATIONALITY MAPPING
  // ============================================================
  // The registry uses numeric option values for country selects.
  // We build a lookup from the live DOM so we don't hardcode values.

  let countryLookup = null;

  function buildCountryLookup() {
    if (countryLookup) return countryLookup;

    // Find any country select on the page to extract options
    const sel = document.querySelector('[id^="principalplaceofbusinesscorpaddresscountryid_"]') ||
                document.querySelector('[id^="corpofficercorpaddresscountryid_"]') ||
                document.querySelector('[id^="corpofficernationalityid_"]');

    if (!sel) {
      log.warn('No country select found — cannot build lookup');
      return {};
    }

    countryLookup = {};
    Array.from(sel.options).forEach(o => {
      if (o.value && o.value !== '0' && o.value !== '') {
        // Map both exact text and lowercase for flexible matching
        countryLookup[o.text.trim().toLowerCase()] = o.value;
      }
    });

    log.info(`Built country lookup: ${Object.keys(countryLookup).length} entries`);
    return countryLookup;
  }

  /**
   * Resolve a country/nationality name to its select option value.
   * Returns the numeric value string, or null if not found.
   */
  function resolveCountry(name) {
    if (!name) return null;
    const lookup = buildCountryLookup();
    const key = name.trim().toLowerCase();

    // Exact match
    if (lookup[key]) return lookup[key];

    // Common aliases
    const aliases = {
      'kiribati': 'kiribati',
      'i-kiribati': 'kiribati',
      'usa': 'united states',
      'us': 'united states',
      'uk': 'united kingdom',
      'nz': 'new zealand',
      'png': 'papua new guinea',
      'aus': 'australia',
    };
    if (aliases[key] && lookup[aliases[key]]) return lookup[aliases[key]];

    // Partial match
    const partial = Object.keys(lookup).find(k => k.includes(key) || key.includes(k));
    if (partial) return lookup[partial];

    log.warn(`Country not resolved: "${name}"`);
    return null;
  }

  // ============================================================
  // ACTIVITY MAPPING
  // ============================================================
  // Form builder checkbox labels → registry select option values

  const ACTIVITY_MAP = {
    'accommodation services': '1',
    'agriculture and livestock': '2',
    'communication, information and it services': '3',
    'construction': '4',
    'education and health': '5',
    'finance and insurance': '6',
    'fishing': '7',
    'food and beverage services': '8',
    'manufacturing': '9',
    'services to businesses': '10',
    'services to households': '11',
    'tourism (excluding accommodation)': '12',
    'transport and storage': '13',
    'wholesale and retail trade': '14',
    'other': '15',
  };

  function resolveActivity(label) {
    if (!label) return null;
    const key = label.trim().toLowerCase();
    if (ACTIVITY_MAP[key]) return ACTIVITY_MAP[key];
    // Partial match
    const match = Object.keys(ACTIVITY_MAP).find(k => k.includes(key) || key.includes(k));
    return match ? ACTIVITY_MAP[match] : null;
  }

  // ============================================================
  // BN-0 / BN-1 SIMPLE FIELD FILLER
  // ============================================================

  function fillBnSimpleFields(data, formInfo) {
    log.info('── Filling simple fields ──');

    // Tab 1: Business Name
    navigateToTab('business-name');

    // For BN-0 re-registrations, copy existing name into "New business name"
    // (the business is re-registering under the same name)
    if (formInfo.formType === 'BN-0' && formInfo.entityName) {
      const result = fieldSetter.setText('corpname', formInfo.entityName);
      log.field('New business name', result === 'ok' ? 'ok' : 'fail',
        result === 'ok' ? `Copied: "${formInfo.entityName}"` : result);
    }

    // Foreign investment radio: JSON "Yes"→"true", "No"→"false"
    if (data.foreignInvestment) {
      const val = data.foreignInvestment === 'Yes' ? 'true' : 'false';
      const result = fieldSetter.setRadio('corpisforeign', val);
      log.field('corpisforeign', result === 'ok' ? 'ok' : 'fail',
        result === 'ok' ? `Set to ${val} (${data.foreignInvestment})` : result);
    }

    // Tab 4: Addresses
    navigateToTab('addresses');

    const addr = data.businessAddress || {};
    const addrFields = [
      ['principalplaceofbusinesscorpaddressstreet1', addr.addr1, 'Address line 1'],
      ['principalplaceofbusinesscorpaddressstreet2', addr.addr2, 'Address line 2'],
      ['principalplaceofbusinesscorpaddressstate',   addr.island, 'Island'],
      ['principalplaceofbusinesscorpaddresscity',    addr.city, 'City/Town'],
      ['principalplaceofbusinesscorpaddresszip',     addr.postcode, 'Postal code'],
    ];

    addrFields.forEach(([field, value, label]) => {
      const result = fieldSetter.setText(field, value);
      if (result === 'skip') {
        log.field(label, 'skip', 'No value in JSON');
      } else if (result === 'ok') {
        log.field(label, 'ok', `"${value}"`);
      } else {
        log.field(label, 'fail', result);
      }
    });

    // Country select (BN address — defaults to Kiribati but we set explicitly)
    // The form builder doesn't collect country for BN business address,
    // but if it's present we'll set it. Otherwise leave as default (Kiribati).
    if (addr.country) {
      const countryVal = resolveCountry(addr.country);
      if (countryVal) {
        const result = fieldSetter.setSelect('principalplaceofbusinesscorpaddresscountryid', countryVal);
        log.field('Country', result === 'ok' ? 'ok' : 'fail', result === 'ok' ? addr.country : result);
      }
    }

    // Address change type radio — set to "1" (Yes/Added) for re-registrations
    {
      const result = fieldSetter.setRadio('principalplaceofbusinesscorpaddresschangetypeid', '1');
      log.field('Address change type', result === 'ok' ? 'ok' : 'fail',
        result === 'ok' ? 'Set to 1 (Added)' : result);
    }

    // Email
    if (addr.email) {
      const result = fieldSetter.setText('corpemailaddress', addr.email);
      log.field('Email', result === 'ok' ? 'ok' : 'fail', result === 'ok' ? addr.email : result);
    }

    // Website (BN-only field)
    if (addr.website) {
      const result = fieldSetter.setText('corpwebsite', addr.website);
      log.field('Website', result === 'ok' ? 'ok' : 'fail', result === 'ok' ? addr.website : result);
    }

    log.info('── Simple fields complete ──');
  }

  // ============================================================
  // GRID CONTROLLER — MutationObserver-based row creation
  // ============================================================
  // This is the foundation for Phase 2. We expose the mechanism
  // here and use it for activities (simplest grid) as a proof.

  /**
   * Click an Add button and wait for the new row to appear in <tbody>.
   * Returns a Promise that resolves to the new <tr> element.
   */
  function addGridRow(addButton, tbody) {
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        observer.disconnect();
        reject(new Error('Timeout waiting for new grid row (3s)'));
      }, 3000);

      const observer = new MutationObserver((mutations) => {
        for (const mut of mutations) {
          for (const node of mut.addedNodes) {
            if (node.nodeName === 'TR' && node.id && node.id.startsWith('row_')) {
              observer.disconnect();
              clearTimeout(timeout);
              // Give Vue a tick to render the fields inside the row
              setTimeout(() => resolve(node), 150);
              return;
            }
          }
        }
      });

      observer.observe(tbody, { childList: true });
      addButton.click();
    });
  }

  /**
   * Click the OK button inside a grid row to commit the entry and close
   * the edit panel. The OK button is a btn-success with text "Ok" or "OK"
   * sitting inside the row's <tr>.
   */
  function confirmGridRow(row) {
    // Find OK button inside this row — text is "Ok" or "OK", class btn-success
    const buttons = row.querySelectorAll('.btn-success');
    for (const btn of buttons) {
      const text = btn.textContent.trim().toLowerCase();
      if (text === 'ok') {
        btn.click();
        log.info(`  Clicked OK to confirm row ${row.id}`);
        return true;
      }
    }
    // Fallback: the OK might be in a nested element within the row's parent
    // (some grids place it outside the <tr> but near it)
    const parent = row.parentElement;
    if (parent) {
      const nearbyBtns = parent.querySelectorAll('.btn-success');
      for (const btn of nearbyBtns) {
        const text = btn.textContent.trim().toLowerCase();
        if (text === 'ok') {
          btn.click();
          log.info(`  Clicked OK (parent scope) to confirm row`);
          return true;
        }
      }
    }
    log.warn(`  No OK button found in row ${row.id}`);
    return false;
  }

  /**
   * Find the Add button and tbody for a grid within a given tab panel.
   * `gridIndex` is 0-based — which table within the tab (owners has 3).
   * 
   * Strategy: getElementById fails on this platform (likely Vue tab component
   * interference). Instead, find ALL table.table-striped elements globally,
   * filter to those inside the target panel using closest(), then index.
   */
  function getGridControls(tabId, gridIndex = 0) {
    // Find all grid tables in the document
    const allTables = document.querySelectorAll('table.table-striped');

    // Filter to tables inside the target tab panel
    // Note: Vue tab component uses #hash as literal ID (e.g. id="#owners" not id="owners")
    const panelTables = Array.from(allTables).filter(t => {
      const panel = t.closest('.tabs-component-panel');
      if (!panel) return false;
      const panelId = panel.id || '';
      return panelId === tabId || panelId === '#' + tabId;
    });

    log.info(`getGridControls: found ${panelTables.length} tables in #${tabId} (need index ${gridIndex})`);

    if (!panelTables[gridIndex]) {
      log.warn(`Grid table index ${gridIndex} not found in #${tabId}`);
      return null;
    }

    const table = panelTables[gridIndex];
    const tbody = table.querySelector('tbody');
    const addBtn = table.querySelector('tfoot .btn-success') ||
                   table.querySelector('tfoot button.btn-success') ||
                   table.querySelector('tfoot button');

    if (!tbody || !addBtn) {
      log.warn(`Grid ${gridIndex} in #${tabId}: tbody=${!!tbody}, addBtn=${!!addBtn}`);
      return null;
    }
    return { table, tbody, addBtn };
  }

  // ============================================================
  // ACTIVITY GRID FILLER
  // ============================================================

  async function fillActivities(activities) {
    if (!activities || activities.length === 0) {
      log.info('No activities to fill');
      return;
    }

    log.info(`── Filling ${activities.length} activities ──`);
    navigateToTab('business-activities');

    // Small delay for tab render
    await sleep(300);

    const grid = getGridControls('business-activities', 0);
    if (!grid) {
      log.error('Activity grid not found!');
      return;
    }

    for (let i = 0; i < activities.length; i++) {
      const actLabel = activities[i];
      const actValue = resolveActivity(actLabel);

      if (!actValue) {
        log.field(`Activity ${i + 1}`, 'fail', `Cannot resolve: "${actLabel}"`);
        continue;
      }

      try {
        log.info(`Adding activity ${i + 1}/${activities.length}: "${actLabel}"`);
        const row = await addGridRow(grid.addBtn, grid.tbody);

        // Activity select — raw UUID id, no field name prefix
        // Target: the <select> inside this row
        const actSelect = row.querySelector('select');
        if (actSelect) {
          actSelect.value = actValue;
          actSelect.dispatchEvent(new Event('change', { bubbles: true }));
          actSelect.dispatchEvent(new Event('input', { bubbles: true }));
          fieldSetter._setVue(actSelect, actValue);
          log.field(`Activity ${i + 1}`, 'ok', `"${actLabel}" → value ${actValue}`);
        } else {
          log.field(`Activity ${i + 1}`, 'fail', 'Select element not found in new row');
        }

        // Change type radio → "1" (Added)
        const changeResult = fieldSetter.setRadio('fakepropertiescorpactivitychangetypeid', '1', row);
        if (changeResult !== 'ok') {
          log.field(`Activity ${i + 1} change type`, 'fail', changeResult);
        }

        // Commencement date — if present in data
        // For BN-0 re-registration, this is typically not collected
        // but we support it for BN-1

        // Confirm the row (click OK to close edit panel)
        confirmGridRow(row);

        // Brief pause between rows
        await sleep(500);

      } catch (e) {
        log.field(`Activity ${i + 1}`, 'fail', e.message);
      }
    }

    log.info('── Activities complete ──');
  }

  // ============================================================
  // OWNER GRID FILLER — All 4 BN owner grids
  // ============================================================
  //
  // Grid layout on Owners tab (#owners) — 3 tables in order:
  //   Index 0: Natural Persons        (corpofficercorpofficertypeid = 1025)
  //   Index 1: Registered Companies   (corpofficercorpofficertypeid = 1026)
  //   Index 2: Other Registered Entities (corpofficercorpofficertypeid = 1028)
  //
  // Beneficial Owners tab (#beneficial-owners) — 1 table:
  //   Index 0: Beneficial Owners      (corpofficercorpofficertypeid = 1045)

  /**
   * Helper: set a field within a grid row, logging the result.
   * Returns true if ok, false otherwise.
   */
  function setRowField(row, fieldName, value, label, setter) {
    setter = setter || 'text';
    let result;
    switch (setter) {
      case 'text':
        result = fieldSetter.setText(fieldName, value, row);
        break;
      case 'select':
        result = fieldSetter.setSelect(fieldName, value, row);
        break;
      case 'selectText':
        result = fieldSetter.setSelectByText(fieldName, value, row);
        break;
      case 'radio':
        result = fieldSetter.setRadio(fieldName, value, row);
        break;
      case 'textarea':
        result = fieldSetter.setTextarea(fieldName, value, row);
        break;
    }
    if (result === 'skip') {
      // Don't log skips for empty optional fields — too noisy
      return true;
    } else if (result === 'ok') {
      log.field(label, 'ok', `"${value}"`);
      return true;
    } else {
      log.field(label, 'fail', result);
      return false;
    }
  }

  /**
   * Resolve a country/nationality name to select option value,
   * falling back to setSelectByText if numeric lookup fails.
   * Sets the select and logs the result.
   */
  function setCountryOrNationality(row, fieldName, value, label) {
    if (!value) return true; // skip empty
    const resolved = resolveCountry(value);
    if (resolved) {
      return setRowField(row, fieldName, resolved, label, 'select');
    } else {
      // Fallback: try matching option text directly
      return setRowField(row, fieldName, value, label, 'selectText');
    }
  }

  /**
   * Fill a single Natural Person owner row (18 fields).
   * Grid A: corpofficercorpofficertypeid = 1025
   */
  function fillNaturalPersonRow(row, owner, index) {
    const lbl = (f) => `Owner ${index} (person): ${f}`;

    // Change type → "1" (Added)
    setRowField(row, 'corpofficerchangetypeid', '1', lbl('change type'), 'radio');

    // Appointment date (Date became owner)
    if (owner.date) {
      setRowField(row, 'corpofficerdateofappointment', owner.date, lbl('appointment date'));
    }

    // Name fields
    setRowField(row, 'corpofficerfirstname', owner.first, lbl('first name'));
    setRowField(row, 'corpofficermiddlename', owner.middle, lbl('middle name'));
    setRowField(row, 'corpofficerlastname', owner.last, lbl('last name'));

    // Name change radio → "false" (No) — default for re-registration
    setRowField(row, 'fakepropertiesisofficernamechanging', 'false', lbl('name change'), 'radio');

    // Nationality
    setCountryOrNationality(row, 'corpofficernationalityid', owner.nationality, lbl('nationality'));
    if (owner.nationality2) {
      setCountryOrNationality(row, 'corpofficerothernationalityid', owner.nationality2, lbl('other nationality'));
    }

    // Gender: form builder has "Male"/"Female", registry wants "1"/"2"
    if (owner.gender) {
      const genderVal = owner.gender.toLowerCase() === 'male' ? '1' :
                        owner.gender.toLowerCase() === 'female' ? '2' : owner.gender;
      setRowField(row, 'corpofficergenderid', genderVal, lbl('gender'), 'radio');
    }

    // DOB: form builder has dobMm + dobYy, registry wants MM/YYYY in one field
    if (owner.dobMm || owner.dobYy) {
      const dob = `${(owner.dobMm || '').padStart(2, '0')}/${owner.dobYy || ''}`;
      setRowField(row, 'corpofficerbirthdate', dob, lbl('DOB'));
    }

    // Address
    setRowField(row, 'corpofficercorpaddressstreet1', owner.addr1, lbl('address line 1'));
    setRowField(row, 'corpofficercorpaddressstreet2', owner.addr2, lbl('address line 2'));
    setCountryOrNationality(row, 'corpofficercorpaddresscountryid', owner.country, lbl('country'));
    setRowField(row, 'corpofficercorpaddressstate', owner.island, lbl('island'));
    setRowField(row, 'corpofficercorpaddresscity', owner.city, lbl('city'));
    setRowField(row, 'corpofficercorpaddresszip', owner.postcode, lbl('postal code'));
  }

  /**
   * Fill a single Registered Company owner row (2 real fields).
   * Grid B: corpofficercorpofficertypeid = 1026
   */
  function fillKiCompanyRow(row, owner, index) {
    const lbl = (f) => `Owner ${index} (ki_company): ${f}`;

    // Change type → "1" (Added)
    setRowField(row, 'corpofficerchangetypeid', '1', lbl('change type'), 'radio');

    // Appointment date
    if (owner.date) {
      setRowField(row, 'corpofficerdateofappointment', owner.date, lbl('appointment date'));
    }

    // Registration number + entity name
    setRowField(row, 'corpofficerregistrationnumber', owner.entityRegNo, lbl('reg number'));
    setRowField(row, 'corpofficerfullname', owner.entityName, lbl('entity name'));
  }

  /**
   * Fill a single Other Registered Entity owner row (14 fields).
   * Grid C: corpofficercorpofficertypeid = 1028
   */
  function fillOtherEntityRow(row, owner, index) {
    const lbl = (f) => `Owner ${index} (other_entity): ${f}`;

    // Change type → "1" (Added)
    setRowField(row, 'corpofficerchangetypeid', '1', lbl('change type'), 'radio');

    // Appointment date
    if (owner.date) {
      setRowField(row, 'corpofficerdateofappointment', owner.date, lbl('appointment date'));
    }

    // Entity details
    setRowField(row, 'corpofficerfullname', owner.oeName, lbl('entity name'));
    setRowField(row, 'corpofficerregistrationnumber', owner.oeRegNo, lbl('reg number'));
    setRowField(row, 'corpofficerentitytype', owner.oeType, lbl('entity type'));

    // Jurisdiction of incorporation (select)
    if (owner.oeJurisdiction) {
      setCountryOrNationality(row, 'corpofficerjurisdictionofincorporationid', owner.oeJurisdiction, lbl('jurisdiction'));
    }

    // Address
    setRowField(row, 'corpofficercorpaddressstreet1', owner.addr1, lbl('address line 1'));
    setRowField(row, 'corpofficercorpaddressstreet2', owner.addr2, lbl('address line 2'));
    setCountryOrNationality(row, 'corpofficercorpaddresscountryid', owner.country, lbl('country'));
    setRowField(row, 'corpofficercorpaddressstate', owner.island, lbl('island'));
    setRowField(row, 'corpofficercorpaddresscity', owner.city, lbl('city'));
    setRowField(row, 'corpofficercorpaddresszip', owner.postcode, lbl('postal code'));
  }

  /**
   * Fill a single Beneficial Owner row (17 fields).
   * Grid D: corpofficercorpofficertypeid = 1045
   * Similar to Natural Person but no gender/DOB, adds relationship textarea.
   */
  function fillBeneficialOwnerRow(row, bo, index) {
    const lbl = (f) => `Beneficial owner ${index}: ${f}`;

    // Change type → "1" (Added)
    setRowField(row, 'corpofficerchangetypeid', '1', lbl('change type'), 'radio');

    // Appointment date
    if (bo.date) {
      setRowField(row, 'corpofficerdateofappointment', bo.date, lbl('appointment date'));
    }

    // Name
    setRowField(row, 'corpofficerfirstname', bo.first, lbl('first name'));
    setRowField(row, 'corpofficermiddlename', bo.middle, lbl('middle name'));
    setRowField(row, 'corpofficerlastname', bo.last, lbl('last name'));

    // Name change radio → "false"
    setRowField(row, 'fakepropertiesisofficernamechanging', 'false', lbl('name change'), 'radio');

    // Nationality
    setCountryOrNationality(row, 'corpofficernationalityid', bo.nationality, lbl('nationality'));
    if (bo.nationality2) {
      setCountryOrNationality(row, 'corpofficerothernationalityid', bo.nationality2, lbl('other nationality'));
    }

    // Address
    setRowField(row, 'corpofficercorpaddressstreet1', bo.addr1, lbl('address line 1'));
    setRowField(row, 'corpofficercorpaddressstreet2', bo.addr2, lbl('address line 2'));
    setCountryOrNationality(row, 'corpofficercorpaddresscountryid', bo.country, lbl('country'));
    setRowField(row, 'corpofficercorpaddressstate', bo.island, lbl('island'));
    setRowField(row, 'corpofficercorpaddresscity', bo.city, lbl('city'));
    setRowField(row, 'corpofficercorpaddresszip', bo.postcode, lbl('postal code'));

    // Relationship textarea — UNIQUE: no UUID suffix, direct ID
    // The briefing says id is just "corpofficerdescriptionofrelationship"
    if (bo.relationship) {
      const ta = row.querySelector('textarea') ||
                 row.querySelector('[id="corpofficerdescriptionofrelationship"]') ||
                 row.querySelector('[id^="corpofficerdescriptionofrelationship"]');
      if (ta) {
        const nativeSetter = Object.getOwnPropertyDescriptor(
          HTMLTextAreaElement.prototype, 'value'
        )?.set;
        if (nativeSetter) nativeSetter.call(ta, bo.relationship);
        else ta.value = bo.relationship;
        ta.dispatchEvent(new Event('input', { bubbles: true }));
        ta.dispatchEvent(new Event('change', { bubbles: true }));
        ta.dispatchEvent(new Event('blur', { bubbles: true }));
        fieldSetter._setVue(ta, bo.relationship);
        log.field(lbl('relationship'), 'ok', `"${bo.relationship}"`);
      } else {
        log.field(lbl('relationship'), 'fail', 'Textarea not found in row');
      }
    }
  }

  /**
   * Process a group of owners against a specific grid.
   * Clicks Add for each, waits for the new row, then fills it.
   */
  async function fillOwnerGrid(tabId, gridIndex, owners, fillRowFn, gridLabel) {
    if (!owners || owners.length === 0) return;

    log.info(`── Filling ${owners.length} ${gridLabel} ──`);

    const grid = getGridControls(tabId, gridIndex);
    if (!grid) {
      log.error(`${gridLabel} grid not found (tab: ${tabId}, index: ${gridIndex})`);
      return;
    }

    for (let i = 0; i < owners.length; i++) {
      try {
        log.info(`Adding ${gridLabel} ${i + 1}/${owners.length}`);
        const row = await addGridRow(grid.addBtn, grid.tbody);
        fillRowFn(row, owners[i], i + 1);
        confirmGridRow(row);
        // Pause between rows for Vue to settle
        await sleep(500);
      } catch (e) {
        log.field(`${gridLabel} ${i + 1}`, 'fail', e.message);
      }
    }

    log.info(`── ${gridLabel} complete ──`);
  }

  /**
   * Main owner fill orchestrator.
   * Groups owners by type, processes each group against its grid.
   * 
   * IMPORTANT: We snapshot all grid references BEFORE adding any rows,
   * because adding rows to one grid can change the table count within
   * the panel (Vue re-renders can create/merge tables).
   */
  async function fillOwners(owners) {
    if (!owners || owners.length === 0) {
      log.info('No owners to fill');
      return;
    }

    // Group by type
    const natural = owners.filter(o => o.type === 'person');
    const kiCo = owners.filter(o => o.type === 'ki_company');
    const other = owners.filter(o => o.type === 'other_entity');

    log.info(`── Owner fill: ${owners.length} total — ${natural.length} persons, ${kiCo.length} ki companies, ${other.length} other entities ──`);

    // Navigate to Owners tab
    navigateToTab('owners');
    await sleep(400);

    // Snapshot all 3 grid references NOW before any rows are added
    const grid0 = getGridControls('owners', 0); // Natural Persons
    const grid1 = getGridControls('owners', 1); // Ki Companies
    const grid2 = getGridControls('owners', 2); // Other Entities

    // Fill Natural Persons (grid 0)
    if (natural.length > 0) {
      if (grid0) {
        log.info(`── Filling ${natural.length} natural person owners ──`);
        for (let i = 0; i < natural.length; i++) {
          try {
            log.info(`Adding natural person owners ${i + 1}/${natural.length}`);
            const row = await addGridRow(grid0.addBtn, grid0.tbody);
            fillNaturalPersonRow(row, natural[i], i + 1);
            confirmGridRow(row);
            await sleep(500);
          } catch (e) {
            log.field(`natural person ${i + 1}`, 'fail', e.message);
          }
        }
        log.info('── natural person owners complete ──');
      } else {
        log.error('Natural persons grid not found (index 0)');
      }
    }

    // Fill Ki Companies (grid 1)
    if (kiCo.length > 0) {
      if (grid1) {
        log.info(`── Filling ${kiCo.length} Ki company owners ──`);
        for (let i = 0; i < kiCo.length; i++) {
          try {
            log.info(`Adding Ki company owners ${i + 1}/${kiCo.length}`);
            const row = await addGridRow(grid1.addBtn, grid1.tbody);
            fillKiCompanyRow(row, kiCo[i], i + 1);
            confirmGridRow(row);
            await sleep(500);
          } catch (e) {
            log.field(`ki company ${i + 1}`, 'fail', e.message);
          }
        }
        log.info('── Ki company owners complete ──');
      } else {
        log.error('Ki company grid not found (index 1)');
      }
    }

    // Fill Other Entities (grid 2)
    if (other.length > 0) {
      if (grid2) {
        log.info(`── Filling ${other.length} other entity owners ──`);
        for (let i = 0; i < other.length; i++) {
          try {
            log.info(`Adding other entity owners ${i + 1}/${other.length}`);
            const row = await addGridRow(grid2.addBtn, grid2.tbody);
            fillOtherEntityRow(row, other[i], i + 1);
            confirmGridRow(row);
            await sleep(500);
          } catch (e) {
            log.field(`other entity ${i + 1}`, 'fail', e.message);
          }
        }
        log.info('── other entity owners complete ──');
      } else {
        log.error('Other entities grid not found (index 2)');
      }
    }
  }

  /**
   * Beneficial owners — separate tab, single grid.
   */
  async function fillBeneficialOwners(beneficialOwners) {
    if (!beneficialOwners || beneficialOwners.length === 0) {
      log.info('No beneficial owners to fill');
      return;
    }

    navigateToTab('beneficial-owners');
    await sleep(400);

    await fillOwnerGrid('beneficial-owners', 0, beneficialOwners, fillBeneficialOwnerRow, 'beneficial owners');
  }

  // ============================================================
  // JSON PARSER
  // ============================================================

  function parseInput(raw) {
    // Strip DATA_START/DATA_END markers if present
    let json = raw;
    const startMarker = '===DATA_START===';
    const endMarker = '===DATA_END===';
    const startIdx = json.indexOf(startMarker);
    const endIdx = json.indexOf(endMarker);
    if (startIdx !== -1 && endIdx !== -1) {
      json = json.substring(startIdx + startMarker.length, endIdx);
    }
    json = json.trim();

    try {
      return JSON.parse(json);
    } catch (e) {
      throw new Error(`Invalid JSON: ${e.message}\nFirst 200 chars: ${json.substring(0, 200)}`);
    }
  }

  // ============================================================
  // PRE-FILL PREVIEW
  // ============================================================

  function showPreview(data, formInfo) {
    console.log('\n[AutoFill] ══════════════════════════════════');
    console.log('[AutoFill] PRE-FILL PREVIEW');
    console.log('[AutoFill] ══════════════════════════════════');
    console.log(`[AutoFill] Form: ${formInfo.formType}`);
    console.log(`[AutoFill] Entity: ${formInfo.entityName || '(new)'}`);
    console.log(`[AutoFill] JSON form: ${data.formId || data.formName || '?'}`);
    console.log(`[AutoFill] Business name: ${data.businessName || '?'}`);
    console.log(`[AutoFill] Reg#: ${data.regNumber || '(new)'}`);

    if (data.foreignInvestment) {
      console.log(`[AutoFill] Foreign investment: ${data.foreignInvestment}`);
    }

    if (data.owners) {
      const types = {};
      data.owners.forEach(o => { types[o.type] = (types[o.type] || 0) + 1; });
      console.log(`[AutoFill] Owners: ${data.owners.length} total — ${JSON.stringify(types)}`);
    }

    if (data.beneficialOwners && data.beneficialOwners.length) {
      console.log(`[AutoFill] Beneficial owners: ${data.beneficialOwners.length}`);
    }

    if (data.businessAddress) {
      const a = data.businessAddress;
      const parts = [a.addr1, a.addr2, a.island, a.city, a.postcode].filter(Boolean);
      console.log(`[AutoFill] Address: ${parts.join(', ')}`);
      if (a.email) console.log(`[AutoFill] Email: ${a.email}`);
      if (a.website) console.log(`[AutoFill] Website: ${a.website}`);
    }

    if (data.activities) {
      console.log(`[AutoFill] Activities: ${data.activities.length} — ${data.activities.join(', ')}`);
    }

    console.log('[AutoFill] ══════════════════════════════════\n');
  }

  // ============================================================
  // MAIN ENTRY POINT
  // ============================================================

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  async function run() {
    log.info(`${TOOL_NAME} v${VERSION}`);
    log.info('Starting...');
    log.reset();

    // 1. Detect form type
    const formInfo = detectForm();
    if (formInfo.formType === 'unknown') {
      log.error('Could not detect form type. Is this a BN filing form?');
      log.info('Expected corporiginalcorptypeid = "3" for business names.');
      return;
    }

    // 2. Get JSON input
    const raw = prompt(
      `${TOOL_NAME} v${VERSION}\n\n` +
      `Detected: ${formInfo.formType} — ${formInfo.entityName || 'New'}\n\n` +
      `Paste the JSON data from the form builder PDF.\n` +
      `(Content between ===DATA_START=== and ===DATA_END=== markers)`
    );

    if (!raw || !raw.trim()) {
      log.info('Cancelled — no input provided.');
      return;
    }

    // 3. Parse JSON
    let data;
    try {
      data = parseInput(raw);
    } catch (e) {
      log.error(e.message);
      alert('Error: ' + e.message);
      return;
    }

    // 4. Validate form match
    if (formInfo.formType === 'BN-0' && data.formId) {
      const normalId = data.formId.toLowerCase().replace('-', '');
      if (normalId !== 'bn0' && normalId !== 'bn1') {
        log.warn(`Form mismatch: registry is ${formInfo.formType} but JSON formId is "${data.formId}"`);
        if (!confirm(`Warning: This looks like a ${data.formId} form but the registry is showing ${formInfo.formType}. Continue anyway?`)) {
          log.info('Cancelled by user (form mismatch).');
          return;
        }
      }
    }

    // 5. Preview
    showPreview(data, formInfo);

    // 6. Build country lookup early
    buildCountryLookup();

    // 7. Fill simple fields (address, foreign investment)
    fillBnSimpleFields(data, formInfo);

    // 8. Fill activities
    await fillActivities(data.activities);

    // 9. Fill owners (Phase 2 stubs)
    await fillOwners(data.owners);
    await fillBeneficialOwners(data.beneficialOwners);

    // 10. Summary
    const counts = log.summary();

    // Navigate back to first tab for review
    navigateToTab('business-name');

    alert(
      `${TOOL_NAME} complete!\n\n` +
      `✓ ${counts.ok} fields filled\n` +
      `– ${counts.skip} skipped (empty)\n` +
      `✗ ${counts.fail} failed\n\n` +
      `Check the console (F12) for details.\n` +
      `Review all tabs before submitting!`
    );
  }

  // ============================================================
  // VERSION CHECK + ENVIRONMENT VALIDATION
  // ============================================================

  function validateEnvironment() {
    const checks = [];

    // Check we're on the right site
    if (!location.hostname.includes('paradigmapps.com')) {
      checks.push('WARNING: Not on paradigmapps.com — this tool is designed for ki.ua.paradigmapps.com');
    }

    // Check for Vue
    const anyVueEl = document.querySelector('[data-v-494da2ed]');
    if (!anyVueEl) {
      checks.push('WARNING: No Vue scoped elements found — form may not have loaded yet');
    }

    // Check for jQuery
    if (typeof jQuery === 'undefined') {
      checks.push('WARNING: jQuery not found');
    }

    // Check for a filing form
    const filingForm = document.querySelector('[id^="corpfilingcorpfilingtypeid_"]');
    if (!filingForm) {
      checks.push('WARNING: No filing type field found — are you on a filing form?');
    }

    if (checks.length > 0) {
      log.warn('Environment checks:');
      checks.forEach(c => log.warn('  ' + c));
    } else {
      log.info('Environment OK');
    }

    return checks.length === 0;
  }

  // ============================================================
  // BOOT
  // ============================================================

  // Expose for console access
  window.KiAutoFill = {
    version: VERSION,
    run,
    fieldSetter,
    detectForm,
    buildCountryLookup,
    resolveCountry,
    resolveActivity,
    addGridRow,
    getGridControls,
    log,
    // Convenience: fill a single field for testing
    test: {
      setText: (field, val) => fieldSetter.setText(field, val),
      setSelect: (field, val) => fieldSetter.setSelect(field, val),
      setSelectByText: (field, text) => fieldSetter.setSelectByText(field, text),
      setRadio: (field, val) => fieldSetter.setRadio(field, val),
      find: (field) => fieldSetter.find(field),
    }
  };

  validateEnvironment();
  log.info(`Loaded. Type KiAutoFill.run() or it will start automatically.`);

  // Auto-run after a brief delay (gives console users time to see the loaded message)
  setTimeout(run, 500);

})();
