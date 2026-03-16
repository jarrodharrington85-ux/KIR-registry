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
            // Try space-collapsed match (PDF line-wrap artifacts)
            const collapsed = normalised.replace(/[\s\u00A0\u200B]+/g, '');
            const collapsedMatch = Array.from(el.options).find(
              o => o.text.trim().toLowerCase().replace(/[\s\u00A0\u200B]+/g, '') === collapsed
            );
            if (collapsedMatch) {
              el.value = collapsedMatch.value;
            } else {
              return `text_not_found: "${text}"`;
            }
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
    },

    /**
     * Set a custom Vue dropdown component (dropdown-container with selectedItem/itemList).
     * These are NOT native <select> elements — they're Vue components that render
     * an <input class="no-input"> as display, with selection managed via Vue $data.
     * 
     * @param {string} fieldName - field ID prefix (e.g. 'corpofficercorpaddressstate')
     * @param {string} value - text value to select (e.g. 'Tarawa')
     * @param {Element} [container] - optional container to search within
     * @returns {string} 'ok', 'skip', 'not_found', 'no_match', or error
     */
    setVueDropdown(fieldName, value, container) {
      if (!value) return 'skip';
      const el = this.find(fieldName, container);
      if (!el) return 'not_found';

      try {
        // Walk up to find the Vue dropdown component (has selectedItem + itemList)
        let node = el;
        let vm = null;
        for (let i = 0; i < 10; i++) {
          if (node.__vue__ && 'selectedItem' in node.__vue__.$data) {
            vm = node.__vue__;
            break;
          }
          node = node.parentElement;
          if (!node) break;
        }

        if (!vm) {
          // Fallback: no Vue dropdown found, try setText
          return this.setText(fieldName, value, container);
        }

        // Find matching item in itemList (case-insensitive, space-collapsed)
        const normalised = value.trim().toLowerCase();
        const collapsed = normalised.replace(/[\s\u00A0]+/g, '');
        
        let match = null;
        if (vm.itemList && Array.isArray(vm.itemList)) {
          match = vm.itemList.find(item => {
            const itemStr = (typeof item === 'string' ? item : item.text || item.name || '').trim().toLowerCase();
            return itemStr === normalised || itemStr.replace(/[\s\u00A0]+/g, '') === collapsed;
          });
        }

        if (match === null || match === undefined) {
          return `no_match: "${value}" not in itemList (${vm.itemList ? vm.itemList.length : 0} items)`;
        }

        // Use itemPicked method if available — this is the proper component API
        if (typeof vm.itemPicked === 'function') {
          vm.itemPicked(match);
          return 'ok';
        }

        // Fallback: set selectedItem directly
        vm.selectedItem = match;
        if (typeof vm.$emit === 'function') {
          vm.$emit('input', match);
          vm.$emit('change', match);
        }

        return 'ok';
      } catch (e) {
        return 'error: ' + e.message;
      }
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

    // Space-collapsed match (handles PDF line-wrap artifacts like "Kiri bati" → "kiribati")
    // Also strip non-breaking spaces (char 160) and other Unicode whitespace
    const collapsed = key.replace(/[\s\u00A0\u200B]+/g, '');
    const collapsedMatch = Object.keys(lookup).find(k => k.replace(/[\s\u00A0\u200B]+/g, '') === collapsed);
    if (collapsedMatch) return lookup[collapsedMatch];

    // Last resort: try matching just the collapsed key against collapsed lookup keys
    // using normalize to handle any Unicode oddities
    const normalCollapsed = collapsed.normalize('NFC');
    const lastResort = Object.keys(lookup).find(k => 
      k.replace(/[\s\u00A0\u200B]+/g, '').normalize('NFC') === normalCollapsed
    );
    if (lastResort) return lookup[lastResort];

    log.warn(`Country not resolved: "${name}" (key="${key}", collapsed="${key.replace(/[\s\u00A0\u200B]+/g, '')}", charCodes=${Array.from(key).map(c => c.charCodeAt(0)).join(',')})`);
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

  async function fillBnSimpleFields(data, formInfo) {
    log.info('── Filling simple fields ──');

    // Tab 1: Business Name
    navigateToTab('business-name');

    // For BN-0 re-registrations, "New business name" is filled at the END
    // of the script (after all tabs are done) because the entity-name-input
    // component clears its value when navigating away from this tab.
    // See the end of run() for the actual typing.

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

    // Set country — use forceCountryChange to trigger full VFG schema re-evaluation
    // This toggles the country model, which switches the layout (e.g. State→Island, removes City for Kiribati)
    const addrCountry = addr.country || 'Kiribati';
    const countryVal = resolveCountry(addrCountry);
    if (countryVal) {
      const result = fieldSetter.setSelect('principalplaceofbusinesscorpaddresscountryid', countryVal);
      log.field('Country', result === 'ok' ? 'ok' : 'fail', result === 'ok' ? addrCountry : result);
      // Force the VFG schema to re-evaluate by toggling the model
      await forceCountryChange('principalplaceofbusinesscorpaddresscountryid', countryVal, addrCountry.toUpperCase());
    }

    // Now set the remaining address fields
    const addrFields = [
      ['principalplaceofbusinesscorpaddressstreet1', addr.addr1, 'Address line 1'],
      ['principalplaceofbusinesscorpaddressstreet2', addr.addr2, 'Address line 2'],
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

    // Handle island — custom Vue dropdown component (not a native <select>)
    if (addr.island) {
      const result = fieldSetter.setVueDropdown('principalplaceofbusinesscorpaddressstate', addr.island);
      log.field('Island', result === 'ok' ? 'ok' : 'fail', result === 'ok' ? `"${addr.island}"` : result);
    } else {
      log.field('Island', 'skip', 'No value in JSON');
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

    // Resolve commencement date: use original formation date from hidden field
    // (for re-registrations, activities are deemed to have commenced on original registration)
    const formationEl = fieldSetter.find('corporiginalformationdate');
    const commencementDate = formationEl ? formationEl.value : null;
    if (commencementDate) {
      log.info(`Using formation date as commencement: ${commencementDate}`);
    } else {
      log.warn('No formation date found — commencement date will be skipped');
    }

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

        // Commencement date — use original formation date for re-registrations
        // (the business activities are deemed to have commenced on the original registration date)
        if (commencementDate) {
          const dateResult = fieldSetter.setText('corpactivityproposeddateofcommencement', commencementDate, row);
          if (dateResult === 'ok') {
            log.field(`Activity ${i + 1} commencement`, 'ok', commencementDate);
          } else if (dateResult !== 'skip') {
            log.field(`Activity ${i + 1} commencement`, 'fail', dateResult);
          }
        }

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
  async function fillNaturalPersonRow(row, owner, index) {
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

    // Address — set country FIRST, wait for Vue conditional re-render
    // Address — force country change to trigger VFG schema re-evaluation
    const ownerCountryVal = resolveCountry(owner.country || 'Kiribati');
    setCountryOrNationality(row, 'corpofficercorpaddresscountryid', owner.country, lbl('country'));
    if (ownerCountryVal) {
      await forceCountryChange('corpofficercorpaddresscountryid', ownerCountryVal, (owner.country || 'Kiribati').toUpperCase(), row);
    }

    setRowField(row, 'corpofficercorpaddressstreet1', owner.addr1, lbl('address line 1'));
    setRowField(row, 'corpofficercorpaddressstreet2', owner.addr2, lbl('address line 2'));

    // Island — use Vue dropdown itemPicked
    if (owner.island) {
      const result = fieldSetter.setVueDropdown('corpofficercorpaddressstate', owner.island, row);
      log.field(lbl('island'), result === 'ok' ? 'ok' : 'fail', result === 'ok' ? `"${owner.island}"` : result);
    }

    setRowField(row, 'corpofficercorpaddresszip', owner.postcode, lbl('postal code'));
  }

  /**
   * Fill a single Registered Company owner row (2 real fields).
   * Grid B: corpofficercorpofficertypeid = 1026
   */
  async function fillKiCompanyRow(row, owner, index) {
    const lbl = (f) => `Owner ${index} (ki_company): ${f}`;

    // Change type → "1" (Added)
    setRowField(row, 'corpofficerchangetypeid', '1', lbl('change type'), 'radio');

    // Appointment date
    if (owner.date) {
      setRowField(row, 'corpofficerdateofappointment', owner.date, lbl('appointment date'));
    }

    // Registration number — this is a validated input that auto-populates
    // the entity name. Must type character by character like the business name.
    if (owner.entityRegNo) {
      const el = fieldSetter.find('corpofficerregistrationnumber', row);
      if (el) {
        el.focus();
        await sleep(200);

        const regNo = (owner.entityRegNo || '').trim();
        for (let i = 0; i < regNo.length; i++) {
          el.value = regNo.substring(0, i + 1);
          el.dispatchEvent(new InputEvent('input', {
            bubbles: true,
            data: regNo[i],
            inputType: 'insertText'
          }));
          el.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: regNo[i] }));
          el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: regNo[i] }));
          await sleep(30);
        }
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));

        // Wait for the auto-validation to look up the company and populate the name
        await sleep(2000);

        const stuck = el.value && el.value.length > 0;
        log.field(lbl('reg number'), stuck ? 'ok' : 'fail',
          stuck ? `Typed: "${el.value}"` : `Failed — staff must type "${regNo}" manually`);

        // Check if entity name was auto-populated
        const nameEl = fieldSetter.find('corpofficerfullname', row);
        if (nameEl && nameEl.value) {
          log.field(lbl('entity name'), 'ok', `Auto-populated: "${nameEl.value}"`);
        } else {
          log.field(lbl('entity name'), 'fail',
            `Name not auto-populated — registration number "${regNo}" may not exist in registry`);
        }
      } else {
        log.field(lbl('reg number'), 'fail', 'Field not found');
      }
    }
  }

  /**
   * Fill a single Other Registered Entity owner row (14 fields).
   * Grid C: corpofficercorpofficertypeid = 1028
   */
  async function fillOtherEntityRow(row, owner, index) {
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

    // Address — force country change to trigger VFG schema re-evaluation
    const oeCountryVal = resolveCountry(owner.country || 'Kiribati');
    setCountryOrNationality(row, 'corpofficercorpaddresscountryid', owner.country, lbl('country'));
    if (oeCountryVal) {
      await forceCountryChange('corpofficercorpaddresscountryid', oeCountryVal, (owner.country || 'Kiribati').toUpperCase(), row);
    }

    setRowField(row, 'corpofficercorpaddressstreet1', owner.addr1, lbl('address line 1'));
    setRowField(row, 'corpofficercorpaddressstreet2', owner.addr2, lbl('address line 2'));

    if (owner.island) {
      const result = fieldSetter.setVueDropdown('corpofficercorpaddressstate', owner.island, row);
      log.field(lbl('island'), result === 'ok' ? 'ok' : 'fail', result === 'ok' ? `"${owner.island}"` : result);
    }

    setRowField(row, 'corpofficercorpaddresszip', owner.postcode, lbl('postal code'));
  }

  /**
   * Fill a single Beneficial Owner row (17 fields).
   * Grid D: corpofficercorpofficertypeid = 1045
   * Similar to Natural Person but no gender/DOB, adds relationship textarea.
   */
  async function fillBeneficialOwnerRow(row, bo, index) {
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

    // Address — force country change to trigger VFG schema re-evaluation
    const boCountryVal = resolveCountry(bo.country || 'Kiribati');
    setCountryOrNationality(row, 'corpofficercorpaddresscountryid', bo.country, lbl('country'));
    if (boCountryVal) {
      await forceCountryChange('corpofficercorpaddresscountryid', boCountryVal, (bo.country || 'Kiribati').toUpperCase(), row);
    }

    setRowField(row, 'corpofficercorpaddressstreet1', bo.addr1, lbl('address line 1'));
    setRowField(row, 'corpofficercorpaddressstreet2', bo.addr2, lbl('address line 2'));

    if (bo.island) {
      const result = fieldSetter.setVueDropdown('corpofficercorpaddressstate', bo.island, row);
      log.field(lbl('island'), result === 'ok' ? 'ok' : 'fail', result === 'ok' ? `"${bo.island}"` : result);
    }

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
        await fillRowFn(row, owners[i], i + 1);
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

    // Normalize owner types — PDF line wrapping can insert spaces (e.g. "ki_ company" → "ki_company")
    owners.forEach(o => {
      if (o.type) {
        const before = o.type;
        o.type = o.type.replace(/[\s\u00A0\u200B]+/g, '').toLowerCase();
        if (before !== o.type) {
          log.info(`Owner type normalized: "${before}" → "${o.type}" (charCodes: ${Array.from(before).map(c => c.charCodeAt(0)).join(',')})`);
        }
      }
    });

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
            await fillNaturalPersonRow(row, natural[i], i + 1);
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
            await fillKiCompanyRow(row, kiCo[i], i + 1);
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
            await fillOtherEntityRow(row, other[i], i + 1);
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

    // Clean PDF copy-paste artifacts:
    // 1. Strip page headers/footers (ministry header, page numbers, form title)
    const pdfJunkPatterns = [
      /MINISTRY OF TOURISM,?\s*COMMERCE,?\s*INDUSTRY AND COOPERATIVES/gi,
      /business\.gov\.ki/gi,
      /BN-\d+\s*\|\s*[^\n]*/gi,        // Form title lines like "BN-0 | Application for..."
      /Page\s+\d+\s+of\s+\d+/gi,       // Page numbers
      /Structured data[^\n]*/gi,        // Data page heading
      /This page contains[^\n]*/gi,     // Data page instruction
      /To import[^\n]*/gi,              // Data page instruction
    ];
    pdfJunkPatterns.forEach(p => { json = json.replace(p, ''); });

    // 2. Fix newlines injected by PDF text wrapping inside JSON strings.
    //    JSON strings cannot contain literal newlines. Since the JSON is
    //    minified (single line), any newline or control character is a PDF artifact
    //    or a prompt() paste artifact. Strip them all.
    //    Replace all control chars (0x00-0x1F except nothing) with space.
    json = json.replace(/[\x00-\x1F\x7F]/g, ' ');

    // 3. Collapse multiple spaces (from stripped junk + control chars)
    json = json.replace(/\s{2,}/g, ' ');

    json = json.trim();

    try {
      return JSON.parse(json);
    } catch (e) {
      throw new Error(`Invalid JSON: ${e.message}\nFirst 300 chars: ${json.substring(0, 300)}`);
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

  /**
   * Force a country change on a VFG address block by toggling the model.
   * The Paradigm Apps registry uses Vue FormGenerator with schema rules that
   * control which address fields appear based on CountryId. Simply setting
   * the <select> value doesn't trigger re-evaluation. We must:
   *   1. Find the grandparent Vue component that holds the model (CountryId, State, etc.)
   *   2. Toggle CountryId to a different value, call evaluateRules()
   *   3. Toggle back to the target country, call evaluateRules() again
   * This forces the full layout switch (e.g. State→Island, City removed for Kiribati).
   *
   * @param {string} countryFieldName - ID prefix for the country select (e.g. 'corpofficercorpaddresscountryid')
   * @param {string} countryId - target country ID (e.g. '120' for Kiribati)
   * @param {string} countryName - target country name (e.g. 'KIRIBATI')
   * @param {Element} [container] - optional container to search within
   */
  async function forceCountryChange(countryFieldName, countryId, countryName, container) {
    const el = fieldSetter.find(countryFieldName, container);
    if (!el) {
      log.warn('forceCountryChange: country field not found');
      return;
    }

    // Walk up to find Vue instance, then grandparent with model
    let node = el;
    while (node && !node.__vue__) node = node.parentElement;
    if (!node || !node.__vue__) {
      log.warn('forceCountryChange: no Vue instance found');
      return;
    }

    const gp = node.__vue__.$parent?.$parent;
    if (!gp || !gp.model || !('CountryId' in gp.model)) {
      log.warn('forceCountryChange: grandparent model not found');
      return;
    }

    // Toggle to a dummy country (10 = Andorra) to force schema re-evaluation
    const dummyId = countryId === '10' ? '20' : '10';
    gp.model.CountryId = parseInt(dummyId);
    gp.model.CountryName = 'TEMP';
    if (gp.$parent && typeof gp.$parent.evaluateRules === 'function') {
      gp.$parent.evaluateRules();
    }
    gp.$forceUpdate();
    await sleep(800);

    // Toggle back to target country
    gp.model.CountryId = parseInt(countryId);
    gp.model.CountryName = countryName;
    if (gp.$parent && typeof gp.$parent.evaluateRules === 'function') {
      gp.$parent.evaluateRules();
    }
    gp.$forceUpdate();
    await sleep(800);

    log.info(`Forced country change to ${countryName} (${countryId})`);
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
    await fillBnSimpleFields(data, formInfo);

    // 8. Fill activities
    await fillActivities(data.activities);

    // 9. Fill owners (Phase 2 stubs)
    await fillOwners(data.owners);
    await fillBeneficialOwners(data.beneficialOwners);

    // 10. Fill business name LAST — the entity-name-input component clears
    // its value when navigating away from the tab. By typing it last,
    // we stay on the business-name tab and nothing else triggers a tab change.
    navigateToTab('business-name');
    await sleep(300);

    if (formInfo.formType === 'BN-0') {
      const nameToSet = formInfo.entityName || data.businessName;
      if (nameToSet) {
        const el = fieldSetter.find('corpname');
        if (el) {
          el.focus();
          await sleep(200);

          // Type character by character — the component only accepts real keystroke events
          for (let i = 0; i < nameToSet.length; i++) {
            el.value = nameToSet.substring(0, i + 1);
            el.dispatchEvent(new InputEvent('input', {
              bubbles: true,
              data: nameToSet[i],
              inputType: 'insertText'
            }));
            el.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: nameToSet[i] }));
            el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: nameToSet[i] }));
            await sleep(30);
          }
          el.dispatchEvent(new Event('change', { bubbles: true }));
          await sleep(300);

          const stuck = el.value && el.value.length > 0;
          log.field('New business name', stuck ? 'ok' : 'fail',
            stuck ? `Typed: "${el.value}"` : `Failed — staff must type "${nameToSet}" manually`);
        } else {
          log.field('New business name', 'fail', 'Field not found');
        }
      }
    }

    // 11. Summary
    const counts = log.summary();

    // 12. Show post-fill checklist panel
    showChecklist(data, formInfo, counts);
  }

  // ============================================================
  // POST-FILL CHECKLIST — Docked side panel
  // ============================================================

  function showChecklist(data, formInfo, counts) {
    // Remove any existing panel
    const existing = document.getElementById('autofill-checklist');
    if (existing) existing.remove();

    // Build upload requirements
    const uploads = [];
    const normalType = (t) => (t || '').replace(/\s+/g, '').toLowerCase();
    if (data.owners) {
      data.owners.forEach(o => {
        if (normalType(o.type) === 'person') {
          const name = [o.first, o.middle, o.last].filter(Boolean).join(' ');
          uploads.push(`Government-issued photo ID for <strong>${name}</strong>`);
        }
      });
    }
    if (data.beneficialOwners) {
      data.beneficialOwners.forEach(bo => {
        const name = [bo.first, bo.middle, bo.last].filter(Boolean).join(' ');
        uploads.push(`Government-issued photo ID for <strong>${name}</strong> (beneficial owner)`);
      });
    }
    if (data.foreignInvestment === 'Yes') {
      uploads.push('Kiribati Foreign Investment Registration Certificate');
    }

    // Build checklist items
    const checks = [];

    // Tab reviews
    if (formInfo.formType === 'BN-0') {
      checks.push({
        label: 'Review the <strong>Business Name</strong> tab',
        detail: 'Confirm the name matches the existing registered business name exactly. For re-registration, the name must be identical.'
      });
    } else {
      checks.push({
        label: 'Review the <strong>Business Name</strong> tab',
        detail: 'Confirm the business name is correct. For new registrations, the name must be unique — use the search button to verify availability.'
      });
    }

    if (data.owners && data.owners.length > 0) {
      const ownerNames = data.owners
        .filter(o => normalType(o.type) === 'person')
        .map(o => [o.first, o.last].filter(Boolean).join(' '));
      checks.push({
        label: 'Review the <strong>Owners</strong> tab',
        detail: `${data.owners.length} owner(s) added — check nationality and address${ownerNames.length ? ' for ' + ownerNames.join(', ') : ''}`
      });
    }

    if (data.beneficialOwners && data.beneficialOwners.length > 0) {
      const boNames = data.beneficialOwners.map(bo => [bo.first, bo.last].filter(Boolean).join(' '));
      checks.push({
        label: 'Review the <strong>Beneficial Owners</strong> tab',
        detail: `Check ${boNames.join(', ')} — nationality, address, relationship`
      });
    }

    checks.push({
      label: 'Review the <strong>Addresses</strong> tab',
      detail: 'Check street address and island are correct'
    });

    if (data.activities && data.activities.length > 0) {
      checks.push({
        label: 'Review the <strong>Business Activities</strong> tab',
        detail: `${data.activities.length} activit${data.activities.length === 1 ? 'y' : 'ies'} added — check selections are correct`
      });
    }

    // Build warnings
    const warnings = [];

    if (data.foreignInvestment === 'Yes') {
      warnings.push({
        icon: 'ℹ',
        label: 'Foreign investment is <strong>Yes</strong>',
        detail: 'Confirm certificate is uploaded before submitting',
        type: 'info'
      });
    }

    // Check for pre-formation appointment dates
    const formationEl = fieldSetter.find('corporiginalformationdate');
    const formationDate = formationEl ? formationEl.value : null;
    if (formationDate && data.owners) {
      const hasEarlyDate = data.owners.some(o => {
        if (!o.date) return false;
        // Compare DD/MM/YYYY dates
        const [dd, mm, yyyy] = o.date.split('/');
        const [fdd, fmm, fyyyy] = formationDate.split('/');
        const ownerDate = new Date(yyyy, mm - 1, dd);
        const formDate = new Date(fyyyy, fmm - 1, fdd);
        return ownerDate < formDate;
      });
      if (hasEarlyDate) {
        warnings.push({
          icon: '⚠',
          label: 'Appointment date validation warning',
          detail: 'Some owner dates are before the formation date. This is expected for re-registrations — the registry may show a "dateLessThanMinDate" warning. This can be ignored.',
          type: 'warning'
        });
      }
    }

    // Failed fields warning
    if (counts.fail > 0) {
      const failedList = log._entries
        .filter(e => e.status === 'fail')
        .map(e => `${e.name}: ${e.detail}`)
        .join('<br>');
      warnings.push({
        icon: '✗',
        label: `<strong>${counts.fail} field(s) failed</strong>`,
        detail: `These must be entered manually:<br>${failedList}`,
        type: 'danger'
      });
    }

    // Render panel
    const panel = document.createElement('div');
    panel.id = 'autofill-checklist';
    panel.innerHTML = `
      <style>
        #autofill-checklist {
          position: fixed;
          top: 0;
          right: 0;
          width: 360px;
          height: 100vh;
          background: #fff;
          border-left: 2px solid #d1d5db;
          box-shadow: -4px 0 12px rgba(0,0,0,0.08);
          z-index: 99999;
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
          font-size: 13px;
          color: #1f2937;
          display: flex;
          flex-direction: column;
          overflow: hidden;
        }
        #autofill-checklist * { box-sizing: border-box; margin: 0; padding: 0; }
        .afc-header {
          padding: 16px 20px;
          background: #f8fafc;
          border-bottom: 1px solid #e5e7eb;
          flex-shrink: 0;
        }
        .afc-header h2 { font-size: 16px; font-weight: 700; margin-bottom: 4px; }
        .afc-header p { font-size: 12px; color: #6b7280; }
        .afc-body {
          flex: 1;
          overflow-y: auto;
          padding: 20px;
        }
        .afc-summary {
          display: flex;
          align-items: center;
          gap: 10px;
          padding: 12px 14px;
          background: ${counts.fail > 0 ? '#fef2f2' : '#f0fdf4'};
          border-radius: 6px;
          margin-bottom: 20px;
          border: 1px solid ${counts.fail > 0 ? '#fecaca' : '#bbf7d0'};
        }
        .afc-summary-icon { font-size: 18px; flex-shrink: 0; }
        .afc-summary-text { font-size: 14px; }
        .afc-summary-text strong { font-weight: 600; }
        .afc-section-label {
          font-size: 11px;
          font-weight: 700;
          color: #6b7280;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          margin: 0 0 10px;
        }
        .afc-upload-list {
          background: #fffbeb;
          border: 1px solid #fde68a;
          border-radius: 6px;
          padding: 12px 14px;
          margin-bottom: 20px;
        }
        .afc-upload-list p { font-size: 13px; font-weight: 600; color: #92400e; margin-bottom: 8px; }
        .afc-upload-list ul { padding-left: 18px; font-size: 12px; color: #92400e; line-height: 1.6; }
        .afc-upload-list li { margin-bottom: 3px; }
        .afc-check {
          display: flex;
          gap: 10px;
          align-items: flex-start;
          padding: 10px 0;
          border-bottom: 1px solid #f3f4f6;
        }
        .afc-check:last-child { border-bottom: none; }
        .afc-check-box {
          width: 20px;
          height: 20px;
          border: 1.5px solid #d1d5db;
          border-radius: 4px;
          flex-shrink: 0;
          margin-top: 1px;
          cursor: pointer;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 13px;
          color: transparent;
          background: #fff;
        }
        .afc-check-box.checked {
          background: #059669;
          border-color: #059669;
          color: #fff;
        }
        .afc-check-label { font-size: 13px; line-height: 1.4; }
        .afc-check-detail { font-size: 12px; color: #6b7280; margin-top: 3px; line-height: 1.4; }
        .afc-warning {
          display: flex;
          gap: 10px;
          align-items: flex-start;
          padding: 12px 14px;
          border-radius: 6px;
          margin-bottom: 8px;
        }
        .afc-warning:last-child { margin-bottom: 20px; }
        .afc-warning.info { background: #eff6ff; border: 1px solid #bfdbfe; }
        .afc-warning.info * { color: #1e40af; }
        .afc-warning.warning { background: #fffbeb; border: 1px solid #fde68a; }
        .afc-warning.warning * { color: #92400e; }
        .afc-warning.danger { background: #fef2f2; border: 1px solid #fecaca; }
        .afc-warning.danger * { color: #991b1b; }
        .afc-warning-icon { font-size: 16px; flex-shrink: 0; }
        .afc-warning-label { font-size: 13px; font-weight: 500; }
        .afc-warning-detail { font-size: 12px; margin-top: 4px; line-height: 1.5; }
        .afc-footer {
          padding: 16px 20px;
          border-top: 1px solid #e5e7eb;
          background: #f8fafc;
          flex-shrink: 0;
        }
        .afc-close-btn {
          width: 100%;
          background: #1d4ed8;
          color: #fff;
          border: none;
          border-radius: 6px;
          padding: 12px 20px;
          font-size: 14px;
          font-weight: 600;
          cursor: pointer;
          letter-spacing: 0.01em;
        }
        .afc-close-btn:hover { background: #1e40af; }
        .afc-version { display: block; text-align: center; font-size: 11px; color: #9ca3af; margin-top: 10px; }
      </style>

      <div class="afc-header">
        <h2>Auto-fill complete</h2>
        <p>${data.businessName || formInfo.entityName || 'Business name'} — ${formInfo.formType} ${formInfo.formType === 'BN-0' ? 're-registration' : 'registration'}</p>
      </div>

      <div class="afc-body">
        <div class="afc-summary">
          <span class="afc-summary-icon">${counts.fail > 0 ? '⚠' : '✓'}</span>
          <div class="afc-summary-text">
            <strong>${counts.ok} fields filled</strong>, ${counts.skip} skipped${counts.fail > 0 ? `, <strong>${counts.fail} failed</strong>` : ''}
          </div>
        </div>

        ${uploads.length > 0 ? `
          <div class="afc-section-label">Upload required</div>
          <div class="afc-upload-list">
            <p>⚠ Upload with this filing:</p>
            <ul>${uploads.map(u => `<li>${u}</li>`).join('')}</ul>
          </div>
        ` : ''}

        ${warnings.length > 0 ? `
          <div class="afc-section-label">Notices</div>
          ${warnings.map(w => `
            <div class="afc-warning ${w.type}">
              <span class="afc-warning-icon">${w.icon}</span>
              <div>
                <div class="afc-warning-label">${w.label}</div>
                <div class="afc-warning-detail">${w.detail}</div>
              </div>
            </div>
          `).join('')}
        ` : ''}

        <div class="afc-section-label">Review before submitting</div>
        ${checks.map((c, i) => `
          <div class="afc-check">
            <div class="afc-check-box" data-idx="${i}" onclick="this.classList.toggle('checked'); this.textContent = this.classList.contains('checked') ? '✓' : '';">
            </div>
            <div>
              <div class="afc-check-label">${c.label}</div>
              <div class="afc-check-detail">${c.detail}</div>
            </div>
          </div>
        `).join('')}
      </div>

      <div class="afc-footer">
        <button class="afc-close-btn" onclick="document.getElementById('autofill-checklist').remove();">Close checklist</button>
        <span class="afc-version">Auto-Fill v${VERSION} — review all tabs before clicking Submit</span>
      </div>
    `;

    document.body.appendChild(panel);
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
