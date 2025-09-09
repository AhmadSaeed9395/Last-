import { itemsList } from './items.js';
import { materialsList } from './materials.js';
import { workmanshipList } from './workmanship.js';
import { laborList } from './labor.js';

class ConstructionCalculator {
    // Add number formatting function
    formatNumber(number) {
        if (isNaN(number) || number === null || number === undefined) return '0.00';
        return number.toLocaleString('en-US', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        });
    }

    constructor() {
        this.initializeElements();
        this.loadMainItems();
        this.setupEventListeners();
        this.resourcesList = [...materialsList, ...workmanshipList, ...laborList];
        this.customPrices = new Map(); // Store custom prices
        this.customUnits = new Map(); // Store custom units
        this.loadPricesSection();
        this.resourcesSection = document.getElementById('resourcesSection');
        this.resourcesAccordion = document.getElementById('resourcesAccordion');
        this.resourcesMaterialsHeader = document.getElementById('resourcesMaterialsHeader');
        this.resourcesMaterialsBody = document.getElementById('resourcesMaterialsBody');
        this.resourcesWorkmanshipHeader = document.getElementById('resourcesWorkmanshipHeader');
        this.resourcesWorkmanshipBody = document.getElementById('resourcesWorkmanshipBody');
        this.resourcesLaborHeader = document.getElementById('resourcesLaborHeader');
        this.resourcesLaborBody = document.getElementById('resourcesLaborBody');
        this.resourceTypeFilter = document.getElementById('resourceTypeFilter');

        // Add item counts to tab labels
        const materialsTab = document.querySelector('button[data-tab="materials"]');
        const workmanshipTab = document.querySelector('button[data-tab="workmanship"]');
        const laborTab = document.querySelector('button[data-tab="labor"]');
        if (materialsTab) materialsTab.textContent = `الخامات (${materialsList.length})`;
        if (workmanshipTab) workmanshipTab.textContent = `المصنعيات (${workmanshipList.length})`;
        if (laborTab) laborTab.textContent = `العمالة (${laborList.length})`;
        


        // Project management elements
        this.projectForm = document.getElementById('projectForm');
        this.projectNameInput = document.getElementById('projectName');
        this.projectCodeInput = document.getElementById('projectCode');
        this.projectTypeInput = document.getElementById('projectType');
        this.projectAreaInput = document.getElementById('projectArea');
        this.projectFloorInput = document.getElementById('projectFloor');
        this.createProjectBtn = document.getElementById('createProjectBtn');
        this.projectsList = document.getElementById('projectsList');
        this.currentProjectDisplay = document.getElementById('currentProjectDisplay');
        // Project data
        this.projects = this.loadProjects();
        this.currentProjectId = this.loadCurrentProjectId();
        // Initialize project management UI
        this.setupProjectManagement();
        // ... rest of constructor ...
        // On project change, reload all data
        this.loadProjectData();
        this.customRates = {};
        this.laborExtrasPerFloor = {}; // resourceName -> extra amount per additional floor
        this.laborFloorLevel = 1; // labor-only floor level

        // In constructor, add modal elements
        this.itemDetailsModal = document.getElementById('itemDetailsModal');
        this.itemDetailsTitle = document.getElementById('itemDetailsTitle');
        this.itemDetailsContent = document.getElementById('itemDetailsContent');
        this.closeItemDetailsModal = document.getElementById('closeItemDetailsModal');
        if (this.closeItemDetailsModal) {
            this.closeItemDetailsModal.onclick = () => this.hideItemDetailsModal();
        }
        // Hide modal on outside click
        if (this.itemDetailsModal) {
            this.itemDetailsModal.addEventListener('click', (e) => {
                if (e.target === this.itemDetailsModal) this.hideItemDetailsModal();
            });
        }
        // Collapsible panels logic
        this.setupCollapsiblePanels();

        // Enhance all number inputs globally
        this.enhanceNumberInputs();


        // Export Summary HTML button logic
        const exportSummaryHtmlBtn = document.getElementById('exportSummaryHtmlBtn');
        if (exportSummaryHtmlBtn) {
            exportSummaryHtmlBtn.onclick = () => this.exportSummaryToHtml();
        }

        // Export Resources HTML button logic
        const exportResourcesHtmlBtn = document.getElementById('exportResourcesHtmlBtn');
        if (exportResourcesHtmlBtn) {
            exportResourcesHtmlBtn.onclick = () => this.exportResourcesToHtml();
        }

        // (already bound above)

        // In constructor after loadPricesSection
        this.loadPricesSection();
        // Show initial resources totals
        this.updateResourcesTotals();
        
        // Initialize undo system
        this.undoStack = [];
        this.maxUndoActions = 10; // Keep last 10 actions
        
        // Initialize summary totals
        this.updateSummaryTotal();
        this.updateSummarySellingTotal();
        
        // Ensure all cards have proper sellPrice in dataset
        this.ensureAllCardsHaveSellPrice();
        
        // Setup selection system for summary cards
        this.setupSelectionSystem();
        
        this.summaryFinalTotal = document.getElementById('summaryFinalTotal');
        this.supervisionPercentage = document.getElementById('supervisionPercentage');
        this.lastDeletedCardData = null;
        this.lastDeletedCardElement = null;
        this.unitPriceDisplay = document.getElementById('unitPriceDisplay');
        
        // Add event listener for supervision percentage
        if (this.supervisionPercentage) {
            this.supervisionPercentage.addEventListener('input', () => {
                this.updateSummaryFinalTotal();
            });
            
            // Also add change event for better reliability
            this.supervisionPercentage.addEventListener('change', () => {
                this.updateSummaryFinalTotal();
            });
            
            // Add focus event to ensure the element is properly initialized
            this.supervisionPercentage.addEventListener('focus', () => {
                // Focus event for reliability
            });
        } else {
            // Try to find it again after a short delay
            setTimeout(() => {
                this.supervisionPercentage = document.getElementById('supervisionPercentage');
                if (this.supervisionPercentage) {
                    this.supervisionPercentage.addEventListener('input', () => {
                        this.updateSummaryFinalTotal();
                    });
                    this.supervisionPercentage.addEventListener('change', () => {
                        this.updateSummaryFinalTotal();
                    });
                }
            }, 100);
        }

        // Migrate and normalize old saved projects on startup
        this.migrateLegacyProjects();
        
        // Ensure selection system is set up after everything is loaded
        window.addEventListener('load', () => {
            setTimeout(() => {
                this.setupSelectionSystem();
                this.updateSelectionCount();
            }, 500);
        });
    }

    // Recalculate all summary cards based on current price lists and custom prices, then save
    recalculateAllCardsAndSave() {
        try {
            const cards = Array.from(this.summaryCards.querySelectorAll('.summary-card'));
            cards.forEach(card => {
                const data = card.cardData;
                if (!data) return;
                // Re-derive unit price from current pricing engine
                const matching = itemsList.filter(i => i['Main Item'] === data.mainItem && i['Sub Item'] === data.subItem);
                let totalCost = 0;
                matching.forEach(i => {
                    const rateKey = `${data.mainItem}||${data.subItem}||${i.Resource}`;
                    const itemQuantity = this.customRates && this.customRates[rateKey] !== undefined
                        ? parseFloat(this.customRates[rateKey])
                        : (parseFloat(i['Quantity per Unit']) || 0);
                    const totalQuantity = (parseFloat(data.quantity) || 0) * itemQuantity;
                    const price = this.getResourcePrice(i.Resource, i.Type) || 0;
                    totalCost += totalQuantity * price;
                });
                const quantity = parseFloat(data.quantity) || 0;
                const newUnitPrice = quantity > 0 ? (totalCost / quantity) : 0;
                data.unitPrice = newUnitPrice;
                data.total = newUnitPrice * quantity;
                // Persist back on the DOM element
                card.cardData = data;
                // Update visible fields in the card header
                const unitPriceEls = card.querySelectorAll('.unit-price-value');
                unitPriceEls.forEach(el => el.textContent = this.formatNumber(newUnitPrice));
                
                // Update selling price with tax and risk
                const sellPriceEls = card.querySelectorAll('.sell-price-header, .sell-price-value');
                const taxPercentage = data.taxPercentage !== undefined ? data.taxPercentage : 14;
                const riskPercentage = data.riskPercentage !== undefined ? data.riskPercentage : 0;
                const sellPrice = newUnitPrice * (1 + riskPercentage / 100) * (1 + taxPercentage / 100);
                sellPriceEls.forEach(el => el.textContent = this.formatNumber(sellPrice));
                
                // Update total with new selling price
                const totalEl = card.querySelector('.item-total');
                if (totalEl) totalEl.innerHTML = `الإجمالي: ${this.formatNumber(sellPrice * quantity)} جنيه`;
                
                // Update body total if it exists
                const bodyTotalEl = card.querySelector('.body-total-value');
                if (bodyTotalEl) bodyTotalEl.textContent = this.formatNumber(sellPrice * quantity);
                
                // Update card dataset
                card.dataset.sellPrice = sellPrice;
            });
            // Save and refresh dependent sections
            this.saveProjectItemsFromDOM();
            this.updateSummaryTotal();
            this.updateSummarySellingTotal();
            this.updateSummaryFinalTotal();
            this.updateResourcesSection();
            this.updateResourcesTotals();
        } catch (e) {
            console.error('Error recalculating all cards:', e);
        }
    }

    // One-time migration for legacy projects (missing prices/units/custom mappings)
    migrateLegacyProjects() {
        try {
            if (!this.projects || typeof this.projects !== 'object') return;
            Object.keys(this.projects).forEach(projectId => {
                const p = this.projects[projectId];
                if (!p) return;
                // Ensure prices object exists
                if (!p.prices) p.prices = { customPrices: {}, customUnits: {} };
                if (!p.prices.customPrices) p.prices.customPrices = {};
                if (!p.prices.customUnits) p.prices.customUnits = {};
                // Ensure customRates exists
                if (!p.customRates) p.customRates = {};
                // Ensure labor extras structure
                if (!p.laborExtrasPerFloor) p.laborExtrasPerFloor = {};
                if (!p.laborFloorLevel) p.laborFloorLevel = 1;
            });
            this.saveProjects();
        } catch (e) {
            console.error('Error migrating legacy projects:', e);
        }
    }

    setupCollapsiblePanels() {
        const panels = [
            { header: 'pricesPanelHeader', content: 'pricesPanelContent', toggle: 'pricesPanelHeader' },
            { header: 'inputPanelHeader', content: 'inputPanelContent', toggle: 'inputPanelHeader' },
            { header: 'resourcesPanelHeader', content: 'resourcesPanelContent', toggle: 'resourcesPanelHeader' },
            { header: 'summaryPanelHeader', content: 'summaryPanelContent', toggle: 'summaryPanelHeader' }
        ];
        panels.forEach(({ header, content }) => {
            const headerEl = document.getElementById(header);
            const contentEl = document.getElementById(content);
            if (!headerEl || !contentEl) return;
            const btn = headerEl.querySelector('.collapsible-toggle');
            // Collapse by default
            contentEl.style.display = 'none';
            btn.setAttribute('aria-expanded', 'false');
            btn.textContent = '+';
            // Toggle logic
            const toggle = (e) => {
                e.stopPropagation();
                const expanded = contentEl.style.display === 'block';
                contentEl.style.display = expanded ? 'none' : 'block';
                btn.setAttribute('aria-expanded', expanded ? 'false' : 'true');
                btn.textContent = expanded ? '+' : '–';
            };
            btn.onclick = toggle;
            headerEl.onclick = (e) => {
                if (e.target.classList.contains('collapsible-toggle')) return;
                toggle(e);
            };
        });
    }

    initializeElements() {
        this.mainItemSelect = document.getElementById('mainItemSelect');
        this.subItemSelect = document.getElementById('subItemSelect');
        this.quantityInput = document.getElementById('quantityInput');
        this.wastePercentInput = document.getElementById('wastePercentInput');
        this.operationPercentInput = document.getElementById('operationPercentInput');
        this.totalCostElement = document.getElementById('totalCost');
        this.resultsSection = document.getElementById('resultsSection');
        this.materialsTable = document.getElementById('materialsTable').querySelector('tbody');
        this.workmanshipTable = document.getElementById('workmanshipTable').querySelector('tbody');
        this.laborTable = document.getElementById('laborTable').querySelector('tbody');
        
        // Prices section elements
        this.materialsGrid = document.getElementById('materials-grid');
        this.workmanshipGrid = document.getElementById('workmanship-grid');
        this.laborGrid = document.getElementById('labor-grid');
        this.laborFloorLevelInput = null; // will be set dynamically
        this.laborExtraInputs = {}; // resourceName: input element

        // Accordion elements
        this.materialsAccordion = document.getElementById('materialsAccordion');
        this.workmanshipAccordion = document.getElementById('workmanshipAccordion');
        this.laborAccordion = document.getElementById('laborAccordion');
        this.materialsHeader = document.getElementById('materialsHeader');
        this.workmanshipHeader = document.getElementById('workmanshipHeader');
        this.laborHeader = document.getElementById('laborHeader');
        this.materialsBody = document.getElementById('materialsBody');
        this.workmanshipBody = document.getElementById('workmanshipBody');
        this.laborBody = document.getElementById('laborBody');
        this.materialsDesc = document.getElementById('materialsDesc');
        this.workmanshipDesc = document.getElementById('workmanshipDesc');
        this.laborDesc = document.getElementById('laborDesc');
        this.materialsTotal = document.getElementById('materialsTotal');
        this.workmanshipTotal = document.getElementById('workmanshipTotal');
        this.laborTotal = document.getElementById('laborTotal');
        this.saveItemBtn = document.getElementById('saveItemBtn');
        this.summarySection = document.getElementById('summarySection');
        this.summaryCards = document.getElementById('summaryCards');
        this.summaryUndoBtn = document.getElementById('summaryUndoBtn');
        this.summaryTotal = document.getElementById('summaryTotal');
        this.summarySellingTotal = document.getElementById('summarySellingTotal');
        this.summaryFinalTotal = document.getElementById('summaryFinalTotal');
        this.supervisionPercentage = document.getElementById('supervisionPercentage');
        this.lastDeletedCardData = null;
        this.lastDeletedCardElement = null;
        this.unitPriceDisplay = document.getElementById('unitPriceDisplay');
    }

    loadMainItems() {
        // Get unique main items from itemsList
        const mainItems = [...new Set(itemsList.map(item => item['Main Item']))]
            .filter(name => name !== 'تأسيس سباكة');
        
        // Custom desired order (these appear first in this exact order)
        const desiredOrder = [
            'الهدم',
            'المباني',
            'تأسيس كهرباء',
            'تأسيس صحي',
            'العزل',
            'تأسيس تكييفات',
            'المحارة',
            'جبسوم بورد',
            'بورسلين',
            'رخام',
            'نقاشة'
        ];
        const priorityIndex = new Map(desiredOrder.map((name, idx) => [name, idx]));
        
        // Sort by custom priority first, then Arabic alphabetical for the rest
        mainItems.sort((a, b) => {
            const ra = priorityIndex.has(a) ? priorityIndex.get(a) : Number.POSITIVE_INFINITY;
            const rb = priorityIndex.has(b) ? priorityIndex.get(b) : Number.POSITIVE_INFINITY;
            if (ra !== rb) return ra - rb;
            return a.localeCompare(b, 'ar');
        });
        
        this.mainItemSelect.innerHTML = '<option value="">-- اختر البند الرئيسي --</option>';
        mainItems.forEach(mainItem => {
            const option = document.createElement('option');
            option.value = mainItem;
            option.textContent = mainItem;
            this.mainItemSelect.appendChild(option);
        });
    }

    loadSubItems(mainItem) {
        this.subItemSelect.innerHTML = '<option value="">-- اختر البند الفرعي --</option>';
        this.subItemSelect.disabled = false;

        const subItems = itemsList
            .filter(item => item['Main Item'] === mainItem && mainItem !== 'تأسيس سباكة')
            .map(item => item['Sub Item'])
            .filter((value, index, self) => self.indexOf(value) === index);

        subItems.forEach(subItem => {
            const option = document.createElement('option');
            option.value = subItem;
            option.textContent = subItem;
            this.subItemSelect.appendChild(option);
        });
    }

    loadPricesSection() {
        // Load materials prices
        this.loadResourcePrices(materialsList, this.materialsGrid, 'materials');
        
        // Load workmanship prices
        this.loadResourcePrices(workmanshipList, this.workmanshipGrid, 'workmanship');
        
        // Load labor prices
        this.loadResourcePrices(laborList, this.laborGrid, 'labor');
    }

    loadResourcePrices(resources, container, type) {
        container.innerHTML = '';
        
        if (type === 'materials') {
            this.loadMaterialsBySectors(container);
        } else if (type === 'workmanship') {
            this.loadWorkmanshipBySectors(container);
        } else if (type === 'labor') {
            // Add floor level input at the top
            const floorDiv = document.createElement('div');
            floorDiv.className = 'labor-floor-level-group';
            const initialFloor = this.laborFloorLevel || 1;
            floorDiv.innerHTML = `
                <label for="laborFloorLevelInput">رقم الدور:</label>
                <input type="number" id="laborFloorLevelInput" min="1" step="1" value="${initialFloor}" style="width: 80px; margin-left: 8px;">
            `;
            container.appendChild(floorDiv);
            this.laborFloorLevelInput = floorDiv.querySelector('#laborFloorLevelInput');
            const onFloorChange = () => {
                const prevFloor = this.laborFloorLevel || 1;
                this.laborFloorLevel = parseInt(this.laborFloorLevelInput.value) || 1;
                this.saveProjectLaborFloorLevel();
                this.updateLaborPricesForFloor(prevFloor);
                // Update resources totals ribbon live
                this.updateResourcesTotals();
            };
            this.laborFloorLevelInput.addEventListener('input', onFloorChange);
            this.laborFloorLevelInput.addEventListener('change', onFloorChange);
            this.laborExtraInputs = {};
            this.loadLaborBySectors(container);
            // After rendering, compute prices for current labor floor
            this.updateLaborPricesForFloor(this.laborFloorLevel || 1);
        } else {
            resources.forEach(resource => {
                this.createPriceItem(resource, container, type);
            });
        }

        // Add event listeners to price inputs and unit selects
        container.querySelectorAll('.price-input').forEach(input => {
            // Clear input on focus for easier typing
            input.addEventListener('focus', (e) => {
                e.target.select();
            });
            
            input.addEventListener('input', (e) => {
                const resourceName = e.target.dataset.resource;
                const newPrice = parseFloat(e.target.value.replace(/,/g, '')) || 0;
                const currentUnit = this.customUnits.get(resourceName);
                const resource = this.getResourceInfo(resourceName);
                
                if (resource && currentUnit && currentUnit !== resource.Unit) {
                    // If we're in alternative unit, convert back to default unit for storage
                    const defaultPrice = this.convertFromAltUnitToDefault(resourceName, newPrice, currentUnit, resource.Unit);
                    this.setCustomPrice(resourceName, defaultPrice);
                } else {
                    this.setCustomPrice(resourceName, newPrice);
                }
                
                // If labor item with extra-per-floor, update the baseline to keep user-entered price fixed for current floor
                const isLabor = e.target.dataset.type === 'labor';
                if (isLabor && this.isLaborWithFloorExtra(resourceName)) {
                    const priceInputEl = e.target;
                    const row = priceInputEl.closest('tr');
                    const extraEl = row ? row.querySelector('.extra-per-floor-input') : null;
                    const floorLevel = parseInt(this.laborFloorLevelInput ? this.laborFloorLevelInput.value : '1') || 1;
                    const extra = extraEl ? (parseFloat(extraEl.value) || 0) : 0;
                    // Compute base for floor 1 so that current shown price remains fixed when changing floors
                    const computedBaseForFloor1 = newPrice - extra * (floorLevel - 1);
                    priceInputEl.dataset.base = isNaN(computedBaseForFloor1) ? '0' : String(computedBaseForFloor1);
                }
                
                this.updateUnitOptions(resourceName, this.customPrices.get(resourceName));
                this.calculate(); // Recalculate immediately
                // Update resources totals ribbon live
                this.updateResourcesTotals();
            });
        });

        container.querySelectorAll('.unit-select').forEach(select => {
            select.addEventListener('change', (e) => {
                const resourceName = e.target.dataset.resource;
                const newUnit = e.target.value;
                this.setCustomUnit(resourceName, newUnit);
                
                // Update the price input to show the price for the selected unit
                this.updatePriceForSelectedUnit(resourceName, newUnit);
                
                this.calculate(); // Recalculate immediately
                // Update resources totals ribbon live
                this.updateResourcesTotals();
            });
        });
    }

    loadMaterialsBySectors(container) {
        // Define sections
        const sections = [
            { title: 'خامات أساسية' },
            { title: 'خامات بورسلين' },
            { title: 'خامات عزل' },
            { title: 'خامات نقاشة' },
            { title: 'خامات جبسوم بورد' },
            { title: 'خامات كهرباء' },
            { title: 'خامات صحية' }
            // removed: { title: 'خامات سباكة' }
        ];

        // Create sections
        sections.forEach(section => {
            const sectionElement = this.createSectorSection(section.title);
            container.appendChild(sectionElement);
        });
    }

    loadWorkmanshipBySectors(container) {
        // Define sections
        const sections = [
            { title: 'مصنعيات مدنية' },
            { title: 'مصنعية تأسيس تكييف' },
            { title: 'مصنعية عزل' },
            { title: 'مصنعية بورسلين' },
            { title: 'مصنعية جبسوم بورد' },
            { title: 'مصنعية نقاشة' },
            { title: 'مصنعية كهرباء' },
            { title: 'مصنعية صحية' }
        ];

        // Create sections
        sections.forEach(section => {
            const sectionElement = this.createSectorSection(section.title);
            container.appendChild(sectionElement);
        });
    }

    loadLaborBySectors(container) {
        // Define sections
        const sections = [
            { title: 'معدات' },
            { title: 'عمالة' }
        ];

        // Create sections
        sections.forEach(section => {
            const sectionElement = this.createSectorSection(section.title);
            container.appendChild(sectionElement);
        });
    }

    createSectorSection(title) {
        let count = 0;
        // Determine the type of sector and get the count
        const materialsSectors = ['خامات أساسية', 'خامات بورسلين', 'خامات عزل', 'خامات نقاشة', 'خامات جبسوم بورد', 'خامات كهرباء', 'خامات صحية'];
        const workmanshipSectors = ['مصنعيات مدنية', 'مصنعية تأسيس تكييف', 'مصنعية عزل', 'مصنعية بورسلين', 'مصنعية جبسوم بورد', 'مصنعية نقاشة', 'مصنعية كهرباء', 'مصنعية صحية'];
        const laborSectors = ['معدات', 'عمالة'];
        if (materialsSectors.includes(title)) {
            count = this.getMaterialsForSector(title).length;
        } else if (workmanshipSectors.includes(title)) {
            count = this.getWorkmanshipForSector(title).length;
        } else if (laborSectors.includes(title)) {
            count = this.getLaborForSector(title).length;
        }
        const section = document.createElement('div');
        section.className = 'sector-section';
        section.innerHTML = `
            <div class="sector-header" onclick="this.parentElement.openSector()">
                <h3>${title} <span style='color:#222;font-weight:bold;font-size:1em;'>(${count})</span></h3>
                <div class="sector-toggle">
                    <span class="toggle-icon">▶</span>
                </div>
            </div>
        `;
        // Add sector functionality to the section
        section.openSector = function() {
            this.showSector(title);
        }.bind(this);
        return section;
    }

    showSector(title) {
        // Create searchable interface
        const searchInterface = document.createElement('div');
        searchInterface.className = 'search-interface';
        searchInterface.innerHTML = `
            <div class="search-content">
                <div class="search-header">
                    <h2>${title}</h2>
                    <button class="close-search" onclick="this.closest('.search-interface').remove()">×</button>
                </div>
                <div class="search-controls">
                    <input type="text" class="search-input" placeholder="ابحث عن المادة..." />
                    <button class="clear-search">مسح البحث</button>
                </div>
                <div class="search-results">
                    <table class="price-table">
                        <thead>
                            <tr>
                                <th>المادة</th>
                                <th>الوحدة الافتراضية</th>
                                <th>السعر</th>
                                <th>الوحدة البديلة</th>
                            </tr>
                        </thead>
                        <tbody>
                        </tbody>
                    </table>
                </div>
            </div>
        `;
        
        document.body.appendChild(searchInterface);
        
        // Determine if this is a materials, workmanship, or labor sector
        const materialsSectors = ['خامات أساسية', 'خامات بورسلين', 'خامات عزل', 'خامات نقاشة', 'خامات جبسوم بورد', 'خامات كهرباء', 'خامات صحية'];
        const workmanshipSectors = ['مصنعيات مدنية', 'مصنعية تأسيس تكييف', 'مصنعية عزل', 'مصنعية بورسلين', 'مصنعية جبسوم بورد', 'مصنعية نقاشة', 'مصنعية كهرباء', 'مصنعية صحية'];
        const laborSectors = ['معدات', 'عمالة'];
        
        const isMaterialsSector = materialsSectors.includes(title);
        const isWorkmanshipSector = workmanshipSectors.includes(title);
        const isLaborSector = laborSectors.includes(title);
        
        let materials = [];
        let workmanship = [];
        let labor = [];
        
        if (isMaterialsSector) {
            materials = this.getMaterialsForSector(title);
        } else if (isWorkmanshipSector) {
            workmanship = this.getWorkmanshipForSector(title);
        } else if (isLaborSector) {
            labor = this.getLaborForSector(title);
        }
        
        const tbody = searchInterface.querySelector('.price-table tbody');
        
        // Add materials
        materials.forEach(resource => {
            this.createPriceItem(resource, tbody, 'materials');
        });
        
        // Add workmanship
        workmanship.forEach(resource => {
            this.createPriceItem(resource, tbody, 'workmanship');
        });
        
        // Add labor
        labor.forEach(resource => {
            this.createPriceItem(resource, tbody, 'labor');
        });
        
        // Add search functionality
        this.addSearchFunctionality(searchInterface, materials, workmanship, labor);
        
        // Focus on search input
        setTimeout(() => {
            searchInterface.querySelector('.search-input').focus();
        }, 100);
    }

    addSearchFunctionality(searchInterface, allMaterials, allWorkmanship, allLabor) {
        const searchInput = searchInterface.querySelector('.search-input');
        const clearBtn = searchInterface.querySelector('.clear-search');
        const tbody = searchInterface.querySelector('.price-table tbody');
        
        // Search function
        const performSearch = (searchTerm) => {
            const filteredMaterials = allMaterials.filter(resource => 
                resource.Resource.toLowerCase().includes(searchTerm.toLowerCase())
            );
            
            const filteredWorkmanship = allWorkmanship.filter(resource => 
                resource.Resource.toLowerCase().includes(searchTerm.toLowerCase())
            );
            
            const filteredLabor = allLabor.filter(resource => 
                resource.Resource.toLowerCase().includes(searchTerm.toLowerCase())
            );
            
            // Clear current table
            tbody.innerHTML = '';
            
            // Add filtered materials
            filteredMaterials.forEach(resource => {
                this.createPriceItem(resource, tbody, 'materials');
            });
            
            // Add filtered workmanship
            filteredWorkmanship.forEach(resource => {
                this.createPriceItem(resource, tbody, 'workmanship');
            });
            
            // Add filtered labor
            filteredLabor.forEach(resource => {
                this.createPriceItem(resource, tbody, 'labor');
            });
            
            // Add event listeners to new items
            this.addSearchEventListeners(searchInterface);
        };
        
        // Search input event
        searchInput.addEventListener('input', (e) => {
            performSearch(e.target.value);
        });
        
        // Clear search button
        clearBtn.addEventListener('click', () => {
            searchInput.value = '';
            performSearch('');
            searchInput.focus();
        });
        
        // Add initial event listeners
        this.addSearchEventListeners(searchInterface);
        
        // Close on escape key
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && searchInterface.parentNode) {
                searchInterface.remove();
            }
        });
    }

    addSearchEventListeners(searchInterface) {
        // Add event listeners to price inputs
        searchInterface.querySelectorAll('.price-input').forEach(input => {
            // Clear input on focus for easier typing
            input.addEventListener('focus', (e) => {
                e.target.select();
            });
            
            input.addEventListener('input', (e) => {
                const resourceName = e.target.dataset.resource;
                const newPrice = parseFloat(e.target.value.replace(/,/g, '')) || 0;
                const currentUnit = this.customUnits.get(resourceName);
                const resource = this.getResourceInfo(resourceName);
                
                if (resource && currentUnit && currentUnit !== resource.Unit) {
                    // If we're in alternative unit, convert back to default unit for storage
                    const defaultPrice = this.convertFromAltUnitToDefault(resourceName, newPrice, currentUnit, resource.Unit);
                    this.setCustomPrice(resourceName, defaultPrice);
                } else {
                    this.setCustomPrice(resourceName, newPrice);
                }
                
                // If labor item with extra-per-floor, update baseline
                const isLabor = e.target.dataset.type === 'labor';
                if (isLabor && this.isLaborWithFloorExtra(resourceName)) {
                    const priceInputEl = e.target;
                    const row = priceInputEl.closest('tr');
                    const extraEl = row ? row.querySelector('.extra-per-floor-input') : null;
                    const floorLevel = parseInt(this.laborFloorLevelInput ? this.laborFloorLevelInput.value : '1') || 1;
                    const extra = extraEl ? (parseFloat(extraEl.value) || 0) : 0;
                    const computedBaseForFloor1 = newPrice - extra * (floorLevel - 1);
                    priceInputEl.dataset.base = isNaN(computedBaseForFloor1) ? '0' : String(computedBaseForFloor1);
                }
                
                this.updateUnitOptions(resourceName, this.customPrices.get(resourceName));
                this.calculate(); // Recalculate immediately
                // Update resources totals ribbon live
                this.updateResourcesTotals();
            });
        });

        // Add event listeners to unit selects
        searchInterface.querySelectorAll('.unit-select').forEach(select => {
            select.addEventListener('change', (e) => {
                const resourceName = e.target.dataset.resource;
                const newUnit = e.target.value;
                this.setCustomUnit(resourceName, newUnit);
                
                // Update the price input to show the price for the selected unit
                this.updatePriceForSelectedUnit(resourceName, newUnit);
                
                this.calculate(); // Recalculate immediately
                // Update resources totals ribbon live
                this.updateResourcesTotals();
            });
        });
    }

    getMaterialsForSector(sectorTitle) {
        const sectorMaterials = {
            'خامات أساسية': [
                'أسمنت أسود', 'أسمنت أبيض', 'رمل مونة', 'رمل ردم', 'مادة لاصقة',
                'طوب أحمر 20 10 5', 'طوب طفلي 20 9 5', 'طوب طفلي 24 11 6', 'طوب مصمت دبل 24 11 11',
                'عتبة', 'زوايا', 'شبك'
            ],
            'خامات بورسلين': [
                'بلاط HDF', 'إكسسوارات'
            ],
            'خامات عزل': [
                'أديبوند', 'ألواح ميمبرين', 'برايمر', 'سيكا 107', 'فوم'
            ],
            'خامات نقاشة': [
                'مشمع', 'سيلر حراري', 'سيلر مائي', 'معجون أكريلك', 'معجون دايتون',
                'صنفرة', 'تيب', 'كرتون', 'بلاستيك 7070', 'دهانات بلاستيك'
            ],
            'خامات جبسوم بورد': [
                'عود تراك 6 متر', 'زاوية معدن'
            ],
            'خامات كهرباء': [
                'انتركم', 'بريزة دفاية', 'بريزة عادية', 'بريزة قوي', 'بواط',
                'تكييف', 'تلفاز', 'تليفون', 'ثرموستات تكييف', 'جاكوزي',
                'جرس أو شفاط', 'داتا', 'دفياتير 3 طرف', 'دفياتير 4 طرف',
                'سخان', 'سخان فوري', 'شيش حصيرة', 'صواعد 16 مل', 'صواعد تليفون',
                'صواعد دش', 'صواعد نت', 'لوحة 12 خط', 'لوحة 18 خط', 'لوحة 24 خط',
                'لوحة 36 خط', 'لوحة 48 خط', 'مخرج إضاءة', 'مخرج إضاءة درج', 'مخرج سماعة'
            ],
            'خامات صحية': [
                'مواسير تغذية بولي 1.5 بوصة', 'مواسير تغذية بولي 1 بوصة', 'مواسير تغذية بولي 3/4 بوصة',
                'خزان دفن', 'جسم دفن 1 مخرج', 'جسم دفن شاور 2 مخرج', 'جسم دفن شاور 3 مخرج', 'جسم دفن حوض 2 مخرج',
                'جيت شاور دفن', 'كوع سن داخلي', 'كوع لحام', 'T لحام', 'كرنك', 'طبة إختبار', 'أفيز',
                'جلبة سن داخلي', 'جلبة سن خارجي', 'T سن داخلي', 'محبس دفن', 'تيفلون بكرة', 'نبل', 'وصلة لي',
                'مواسير صرف 4 بوصة', 'مواسير صرف 3 بوصة', 'مواسير صرف 1.5 بوصة', 'مواسير صرف 1 بوصة',
                'كوع مفتوح 45 * 4 بوصة', 'كوع مفتوح 45 * 3 بوصة', 'كوع مفتوح 45 * 1.5 بوصة', 'كوع مفتوح 45 * 1 بوصة',
                'كوع مقفول 90 * 4 بوصة', 'كوع مقفول 90 * 3 بوصة', 'كوع مقفول 90 * 1.5 بوصة', 'كوع مقفول 90 * 1 بوصة',
                'جلبة 4 بوصة', 'جلبة 3 بوصة', 'جلبة 1.5 بوصة', 'جلبة 1 بوصة',
                'بيبة كيسيل 15*15', 'بيبة كيسيل 10*10', 'بيبة كيسيل 6.5*65', 'بيبة كيسيل 6.5*35',
                'محبس زاوية سمارت هوم', 'جلب تطويل ألماني', 'وش نيكل صيني'
            ]
        };
        
        const materialNames = sectorMaterials[sectorTitle] || [];
        return materialsList.filter(resource => materialNames.includes(resource.Resource));
    }

    getWorkmanshipForSector(sectorTitle) {
        const sectorWorkmanship = {
            'مصنعيات مدنية': [
                'مصنعية طوب طفلي 20 9 5', 'مصنعية طوب طفلي 24 11 6', 'مصنعية طوب مصمت دبل 24 11 11', 'مصنعية طوب أحمر 20 10 5',
                'مصنعية نحاتة', 'مصنعية بياض'
            ],
            'مصنعية تأسيس تكييف': [
                'مصنعية 1.5/2.25 HP', 'مصنعية 3/4 HP', 'مصنعية 5 HP', 'مصنعية صاج'
            ],
            'مصنعية عزل': [
                'مصنعية أنسومات', 'مصنعية سيكا 107', 'مصنعية حراري'
            ],
            'مصنعية بورسلين': [
                'مصنعية بورسلين 120*60', 'مصنعية HDF', 'مصنعية وزر'
            ],
            'مصنعية جبسوم بورد': [
                'مصنعية أبيض مسطح', 'مصنعية أخضر مسطح', 'مصنعية أبيض طولي', 'مصنعية أخضر طولي',
                'مصنعية تجاليد أبيض', 'مصنعية تجاليد أخضر', 'مصنعية قواطيع أبيض', 'مصنعية قواطيع أخضر',
                'مصنعية بيوت ستائر و نور', 'مصنعية تراك ماجنتك'
            ],
            'مصنعية نقاشة': [
                'مصنعية تأسيس نقاشة حوائط', 'مصنعية تأسيس نقاشة أسقف', 'مصنعية تشطيب نقاشة'
            ],
            'مصنعية كهرباء': [
                'مصنعية مخرج إضاءة', 'مصنعية مخرج إضاءة درج', 'مصنعية دفياتير 3 طرف', 'مصنعية دفياتير 4 طرف',
                'مصنعية مخرج سماعة', 'مصنعية جرس أو شفاط', 'مصنعية بريزة عادية', 'مصنعية بريزة قوي',
                'مصنعية بريزة دفاية', 'مصنعية جاكوزي', 'مصنعية سخان', 'مصنعية سخان فوري', 'مصنعية تكييف',
                'مصنعية تليفون', 'مصنعية تلفاز', 'مصنعية داتا', 'مصنعية شيش حصيرة', 'مصنعية ثرموستات تكييف',
                'مصنعية انتركم', 'مصنعية لوحة 12 خط', 'مصنعية لوحة 18 خط', 'مصنعية لوحة 24 خط',
                'مصنعية لوحة 36 خط', 'مصنعية لوحة 48 خط', 'مصنعية صواعد 16 مل', 'مصنعية صواعد نت',
                'مصنعية صواعد تليفون', 'مصنعية صواعد دش'
            ],
            'مصنعية صحية': [
                'حمام الماستر', 'حمام الضيوف', 'المطبخ', 'الأوفيس',
                'تأسيس شاور', 'تأسيس حوض', 'تأسيس قعدة عادية', 'تأسيس خزان دفن', 'تأسيس سخان',
                'Mixer عادي بارد ساخن', 'بيبة 15*15', 'بيبة 10*10', 'جريلة 65', 'جريلة 35',
                'محبس دفن', 'صرف تكييف', 'تأسيس غسالة ملابس', 'تأسيس غسالة أطباق', 'محبس زاوية'
            ]
        };
        const workmanshipNames = sectorWorkmanship[sectorTitle] || [];
        return workmanshipList.filter(resource => workmanshipNames.includes(resource.Resource));
    }

    getLaborForSector(sectorTitle) {
        const sectorLabor = {
            'معدات': [
                'عربية رتش', 'هيلتي'
            ],
            'عمالة': [
                'نظافة', 'تشوين', 'تشوين رمل', 'تشوين أسمنت', 'تشوين طوب',
                'تنزيل رتش', 'تشوين بورسلين', 'تشوين مادة لاصقة', 'لياسة'
            ]
        };
        
        const laborNames = sectorLabor[sectorTitle] || [];
        return laborList.filter(resource => laborNames.includes(resource.Resource));
    }

    createPriceItem(resource, container, type) {
        // Handle NaN values properly
        const unitCost = resource['Unit Cost'];
        const defaultPrice = (unitCost && !isNaN(unitCost)) ? unitCost : 0;
        const storedPrice = this.customPrices.get(resource.Resource) || defaultPrice;
        const currentUnit = this.customUnits.get(resource.Resource) || resource.Unit;
        
        // Calculate the price to display based on current unit
        let displayPrice = storedPrice;
        if (currentUnit !== resource.Unit) {
            displayPrice = this.calculateAltUnitPrice(resource.Resource, storedPrice, resource.Unit, currentUnit);
        }
        
        // Create unit options
        const unitOptions = this.createUnitOptions(resource, storedPrice, currentUnit);
        
        // Determine if we should show the unit selector
        const hasAltUnit = resource['Alt Unit'] !== null && resource['Alt Unit'] !== undefined;
        
        const row = document.createElement('tr');
        row.className = 'price-row';
        
        // If labor and resource needs extra per floor, add extra input
        if (type === 'labor' && this.isLaborWithFloorExtra(resource.Resource)) {
            row.innerHTML = `
                <td class="resource-name">${resource.Resource}</td>
                <td class="default-unit">${resource.Unit}</td>
                <td class="price-input-cell">
                    <input type="number" value="${displayPrice}" min="0" step="0.01" placeholder="أدخل السعر" data-resource="${resource.Resource}" data-type="${type}" class="price-input" data-base="${defaultPrice}">
                </td>
                <td class="extra-per-floor-cell">
                    <input type="number" value="" min="0" step="0.01" placeholder="إضافة لكل دور" data-resource="${resource.Resource}" class="extra-per-floor-input">
                </td>
                <td class="unit-select-cell">
                    ${hasAltUnit ? `<select class="unit-select" data-resource="${resource.Resource}" data-type="${type}">${unitOptions}</select>` : '<span class="no-unit">-</span>'}
                </td>
            `;
            container.appendChild(row);
            // Store reference to extra input
            const extraInput = row.querySelector('.extra-per-floor-input');
            this.laborExtraInputs[resource.Resource] = extraInput;
            // Restore saved extra if exists
            if (this.laborExtrasPerFloor && this.laborExtrasPerFloor.hasOwnProperty(resource.Resource)) {
                extraInput.value = String(this.laborExtrasPerFloor[resource.Resource]);
            }
            // Listen for changes: save and recalc
            extraInput.addEventListener('input', (e) => {
                const val = parseFloat(e.target.value);
                this.laborExtrasPerFloor[resource.Resource] = isNaN(val) ? 0 : val;
                this.saveProjectLaborExtras();
                this.updateLaborPricesForFloor();
                // Update resources totals ribbon live
                this.updateResourcesTotals();
            });
            
            // Add event listener for painting price linking (for labor rows too)
            const priceInput = row.querySelector('.price-input');
            if (priceInput) {
                priceInput.addEventListener('input', (e) => {
                    this.handleGypsumPriceChange(resource.Resource, parseFloat(e.target.value));
                    this.handlePaintingPriceChange(resource.Resource, parseFloat(e.target.value));
                    // Sync رمل ردم and رمل مونة prices
                    this.handleSandPriceSync(resource.Resource, parseFloat(e.target.value));
                });
            }
        } else {
            row.innerHTML = `
                <td class="resource-name">${resource.Resource}</td>
                <td class="default-unit">${resource.Unit}</td>
                <td class="price-input-cell">
                    <input type="number" value="${displayPrice}" min="0" step="0.01" placeholder="أدخل السعر" data-resource="${resource.Resource}" data-type="${type}" class="price-input">
                </td>
                <td class="unit-select-cell">
                    ${hasAltUnit ? `<select class="unit-select" data-resource="${resource.Resource}" data-type="${type}">${unitOptions}</select>` : '<span class="item-details">-</span>'}
                </td>
            `;
            container.appendChild(row);
            
            // Add event listener for gypsum board price linking
            const priceInput = row.querySelector('.price-input');
            if (priceInput) {
                priceInput.addEventListener('input', (e) => {
                    this.handleGypsumPriceChange(resource.Resource, parseFloat(e.target.value));
                    this.handlePaintingPriceChange(resource.Resource, parseFloat(e.target.value));
                    // Sync رمل ردم and رمل مونة prices
                    this.handleSandPriceSync(resource.Resource, parseFloat(e.target.value));
                });
            }
        }
    }

    createUnitOptions(resource, currentPrice, currentUnit) {
        const defaultUnit = resource.Unit;
        const altUnit = resource['Alt Unit'];
        
        // If no alternative unit exists, return null to indicate no selector needed
        if (!altUnit) {
            return null;
        }
        
        let options = `<option value="${defaultUnit}" ${currentUnit === defaultUnit ? 'selected' : ''}>${defaultUnit}</option>`;
        options += `<option value="${altUnit}" ${currentUnit === altUnit ? 'selected' : ''}>${altUnit}</option>`;
        
        return options;
    }

    handleGypsumPriceChange(resourceName, newPrice) {
        // Define gypsum board price relationships
        const gypsumRelations = {
            'مصنعية أبيض مسطح': {
                'مصنعية أبيض طولي': 0.7,        // 70% of flat white
                'مصنعية قواطيع أبيض': 2.0,      // Double flat white
                'مصنعية بيوت ستائر و نور': 0.5,  // Half flat white
                'مصنعية تجاليد أبيض': 1.0       // Same as flat white
            },
            'مصنعية أخضر مسطح': {
                'مصنعية أخضر طولي': 0.7,        // 70% of flat green
                'مصنعية قواطيع أخضر': 2.0,      // Double flat green
                'مصنعية بيوت ستائر و نور': 0.5,  // Half flat green
                'مصنعية تجاليد أخضر': 1.0       // Same as flat green
            }
        };
        
        // Check if this resource affects other gypsum prices
        if (gypsumRelations[resourceName]) {
            const relations = gypsumRelations[resourceName];
            
            Object.entries(relations).forEach(([relatedResource, multiplier]) => {
                // Update the related resource price
                this.customPrices.set(relatedResource, newPrice * multiplier);
                
                // Update the display in the UI if it exists
                this.updateGypsumPriceDisplay(relatedResource, newPrice * multiplier);
            });
        }
    }
    
    updateGypsumPriceDisplay(resourceName, newPrice) {
        // Find and update the price input in the UI
        const priceInputs = document.querySelectorAll(`input[data-resource="${resourceName}"]`);
        priceInputs.forEach(input => {
            if (input.classList.contains('price-input')) {
                input.value = newPrice;
            }
        });
    }

    handlePaintingPriceChange(resourceName, newPrice) {
        // Define painting workmanship price relationships
        const paintingRelations = {
            'مصنعية تأسيس نقاشة حوائط': {
                'مصنعية تشطيب نقاشة': 0.5        // 50% of base price
            },
            'مصنعية تأسيس نقاشة أسقف': {
                'مصنعية تشطيب نقاشة': 0.5        // 50% of base price
            }
        };
        
        // Check if this resource affects other painting prices
        if (paintingRelations[resourceName]) {
            const relations = paintingRelations[resourceName];
            
            Object.entries(relations).forEach(([relatedResource, multiplier]) => {
                // Update the related resource price
                this.customPrices.set(relatedResource, newPrice * multiplier);
                
                // Update the display in the UI if it exists
                this.updatePaintingPriceDisplay(relatedResource, newPrice * multiplier);
            });
        }
    }
    
    updatePaintingPriceDisplay(resourceName, newPrice) {
        // Find and update the price input in the UI
        const priceInputs = document.querySelectorAll(`input[data-resource="${resourceName}"]`);
        priceInputs.forEach(input => {
            if (input.classList.contains('price-input')) {
                input.value = newPrice;
            }
        });
    }

    handleSandPriceSync(resourceName, newPrice) {
        // Sync رمل ردم and رمل مونة prices
        if (resourceName === 'رمل ردم') {
            // If رمل ردم price changes, update رمل مونة to match
            this.customPrices.set('رمل مونة', newPrice);
            this.updateSandPriceDisplay('رمل مونة', newPrice);
        } else if (resourceName === 'رمل مونة') {
            // If رمل مونة price changes, update رمل ردم to match
            this.customPrices.set('رمل ردم', newPrice);
            this.updateSandPriceDisplay('رمل ردم', newPrice);
        }
    }

    updateSandPriceDisplay(resourceName, newPrice) {
        // Find and update the price input in the UI for sand materials
        const priceInputs = document.querySelectorAll(`input[data-resource="${resourceName}"]`);
        priceInputs.forEach(input => {
            if (input.classList.contains('price-input')) {
                input.value = newPrice;
            }
        });
    }

    getUnitConversionFactor(resourceName, fromUnit, toUnit) {
        const resource = this.getResourceInfo(resourceName);
        if (!resource) return 1;

        // Define conversion factors for common units
        const conversions = {
            // Cement conversions
            'أسمنت أسود': {
                'شيكارة': { 'طن': 0.05 }, // 20 bags = 1 ton
                'طن': { 'شيكارة': 20 }
            },
            'أسمنت أبيض': {
                'شيكارة': { 'طن': 0.04 }, // 25 bags = 1 ton
                'طن': { 'شيكارة': 25 }
            },
            'مادة لاصقة': {
                'شيكارة': { 'طن': 0.05 }, // 20 bags = 1 ton
                'طن': { 'شيكارة': 20 }
            },
            // Sand conversions
            'رمل مونة': {
                'م3': { 'نقلة': 0.1 }, // Approximate conversion
                'نقلة': { 'م3': 10 }
            },
            'رمل ردم': {
                'م3': { 'نقلة': 0.1 },
                'نقلة': { 'م3': 10 }
            },
            // Brick conversions
            'طوب أحمر 20 10 5': {
                'طوبة': { '1000 طوبة': 0.001 },
                '1000 طوبة': { 'طوبة': 1000 }
            },
            'طوب طفلي 20 9 5': {
                'طوبة': { '1000 طوبة': 0.001 },
                '1000 طوبة': { 'طوبة': 1000 }
            },
            'طوب طفلي 24 11 6': {
                'طوبة': { '1000 طوبة': 0.001 },
                '1000 طوبة': { 'طوبة': 1000 }
            },
            'طوب مصمت دبل 24 11 11': {
                'طوبة': { '1000 طوبة': 0.001 },
                '1000 طوبة': { 'طوبة': 1000 }
            },
            // Paint and sealant conversions
            'سيلر حراري': {
                'لتر': { 'بستلة 20 لتر': 20 },
                'بستلة 20 لتر': { 'لتر': 0.05 }
            },
            'سيلر مائي': {
                'بستلة 9 لتر': { 'لتر': 0.1111111111111111 },
                'لتر': { 'بستلة 9 لتر': 9 }
            },
            'معجون أكريلك': {
                'كيلو': { 'بستلة 15 كيلو': 15 },
                'بستلة 15 كيلو': { 'كيلو': 0.06666666666666667 }
            },
            'معجون دايتون': {
                'كيلو': { 'بستلة 15 كيلو': 15 },
                'بستلة 15 كيلو': { 'كيلو': 0.06666666666666667 }
            },
            'بلاستيك 7070': {
                'لتر': { 'بستلة 9 لتر': 9 },
                'بستلة 9 لتر': { 'لتر': 0.1111111111111111 }
            },
            'دهانات بلاستيك': {
                'لتر': { 'بستلة 9 لتر': 9 },
                'بستلة 9 لتر': { 'لتر': 0.1111111111111111 }
            }
        };

        const resourceConversions = conversions[resourceName];
        if (resourceConversions && resourceConversions[fromUnit] && resourceConversions[fromUnit][toUnit]) {
            return resourceConversions[fromUnit][toUnit];
        }

        return 1; // No conversion needed or available
    }

    convertPrice(resourceName, price, fromUnit, toUnit) {
        if (fromUnit === toUnit) return price;
        
        const conversionFactor = this.getUnitConversionFactor(resourceName, fromUnit, toUnit);
        // Ensure we don't get NaN in conversion
        const safePrice = isNaN(price) ? 0 : price;
        const safeConversionFactor = isNaN(conversionFactor) ? 1 : conversionFactor;
        return safePrice * safeConversionFactor;
    }

    calculateAltUnitPrice(resourceName, defaultPrice, defaultUnit, altUnit) {
        // Specific conversion factors based on user requirements
        const conversions = {
            // Cement: bag price x number of bags per ton
            'أسمنت أسود': {
                'شيكارة': { 'طن': 20 },
                'طن': { 'شيكارة': 0.05 }
            },
            'أسمنت أبيض': {
                'شيكارة': { 'طن': 25 },
                'طن': { 'شيكارة': 0.04 }
            },
            'مادة لاصقة': {
                'شيكارة': { 'طن': 20 },
                'طن': { 'شيكارة': 0.05 }
            },
            // Sand: نقلة = 3 م3 (price scaling)
            'رمل مونة': {
                'م3': { 'نقلة': 3 },
                'نقلة': { 'م3': 0.3333333333333333 }
            },
            'رمل ردم': {
                'م3': { 'نقلة': 3 },
                'نقلة': { 'م3': 0.3333333333333333 }
            },
            // Paint and sealant conversions
            'سيلر حراري': {
                'لتر': { 'بستلة 20 لتر': 20 },
                'بستلة 20 لتر': { 'لتر': 0.05 }
            },
            'سيلر مائي': {
                'بستلة 9 لتر': { 'لتر': 0.1111111111111111 },
                'لتر': { 'بستلة 9 لتر': 9 }
            },
            'معجون أكريلك': {
                'كيلو': { 'بستلة 15 كيلو': 15 },
                'بستلة 15 كيلو': { 'كيلو': 0.06666666666666667 }
            },
            'معجون دايتون': {
                'كيلو': { 'بستلة 15 كيلو': 15 },
                'بستلة 15 كيلو': { 'كيلو': 0.06666666666666667 }
            },
            'بلاستيك 7070': {
                'لتر': { 'بستلة 9 لتر': 9 },
                'بستلة 9 لتر': { 'لتر': 0.1111111111111111 }
            },
            'دهانات بلاستيك': {
                'لتر': { 'بستلة 9 لتر': 9 },
                'بستلة 9 لتر': { 'لتر': 0.1111111111111111 }
            },
            // Brick conversions
            'طوب أحمر 20 10 5': {
                'طوبة': { '1000 طوبة': 1000 },
                '1000 طوبة': { 'طوبة': 0.001 }
            },
            'طوب طفلي 20 9 5': {
                'طوبة': { '1000 طوبة': 1000 },
                '1000 طوبة': { 'طوبة': 0.001 }
            },
            'طوب طفلي 24 11 6': {
                'طوبة': { '1000 طوبة': 1000 },
                '1000 طوبة': { 'طوبة': 0.001 }
            },
            'طوب مصمت دبل 24 11 11': {
                'طوبة': { '1000 طوبة': 1000 },
                '1000 طوبة': { 'طوبة': 0.001 }
            }
        };

        const resourceConversions = conversions[resourceName];
        if (resourceConversions && resourceConversions[defaultUnit] && resourceConversions[defaultUnit][altUnit]) {
            const conversionFactor = resourceConversions[defaultUnit][altUnit];
            // Ensure we don't get NaN in calculation
            const safeDefaultPrice = isNaN(defaultPrice) ? 0 : defaultPrice;
            const result = safeDefaultPrice * conversionFactor;
            

            
            return result;
        }

        // Ensure we don't return NaN
        return isNaN(defaultPrice) ? 0 : defaultPrice; // No conversion available
    }

    convertFromAltUnitToDefault(resourceName, altUnitPrice, altUnit, defaultUnit) {
        // Specific conversion factors for converting from alternative unit back to default unit
        const conversions = {
            // Cement: ton price back to bag price
            'أسمنت أسود': {
                'طن': { 'شيكارة': 0.05 }
            },
            'أسمنت أبيض': {
                'طن': { 'شيكارة': 0.04 }
            },
            'مادة لاصقة': {
                'طن': { 'شيكارة': 0.05 }
            },
            // Sand: if نقلة price is x, then م3 price is x/3
            'رمل مونة': {
                'نقلة': { 'م3': 0.3333333333333333 }
            },
            'رمل ردم': {
                'نقلة': { 'م3': 0.3333333333333333 }
            },
            // Paint and sealant conversions
            'سيلر حراري': {
                'بستلة 20 لتر': { 'لتر': 0.05 }
            },
            'سيلر مائي': {
                'بستلة 9 لتر': { 'لتر': 0.1111111111111111 }
            },
            'معجون أكريلك': {
                'بستلة 15 كيلو': { 'كيلو': 0.06666666666666667 }
            },
            'معجون دايتون': {
                'بستلة 15 كيلو': { 'كيلو': 0.06666666666666667 }
            },
            'بلاستيك 7070': {
                'بستلة 9 لتر': { 'لتر': 0.1111111111111111 }
            },
            'دهانات بلاستيك': {
                'بستلة 9 لتر': { 'لتر': 0.1111111111111111 }
            },
            // Brick conversions
            'طوب أحمر 20 10 5': {
                '1000 طوبة': { 'طوبة': 0.001 }
            },
            'طوب طفلي 20 9 5': {
                '1000 طوبة': { 'طوبة': 0.001 }
            },
            'طوب طفلي 24 11 6': {
                '1000 طوبة': { 'طوبة': 0.001 }
            },
            'طوب مصمت دبل 24 11 11': {
                '1000 طوبة': { 'طوبة': 0.001 }
            }
        };

        const resourceConversions = conversions[resourceName];
        if (resourceConversions && resourceConversions[altUnit] && resourceConversions[altUnit][defaultUnit]) {
            const conversionFactor = resourceConversions[altUnit][defaultUnit];
            // Ensure we don't get NaN in calculation
            const safeAltUnitPrice = isNaN(altUnitPrice) ? 0 : altUnitPrice;
            return safeAltUnitPrice * conversionFactor;
        }

        // Ensure we don't return NaN
        return isNaN(altUnitPrice) ? 0 : altUnitPrice; // No conversion available
    }

    updateUnitOptions(resourceName, newPrice) {
        // Find all unit selects for this resource and update their options
        const unitSelects = document.querySelectorAll(`.unit-select[data-resource="${resourceName}"]`);
        
        unitSelects.forEach(select => {
            const resource = this.getResourceInfo(resourceName);
            if (resource && resource['Alt Unit']) {
                const defaultUnit = resource.Unit;
                const altUnit = resource['Alt Unit'];
                const currentUnit = this.customUnits.get(resourceName) || defaultUnit;
                
                // Calculate alternative unit price
                const altUnitPrice = this.calculateAltUnitPrice(resourceName, newPrice, defaultUnit, altUnit);
                
                // Update the options
                select.innerHTML = `
                    <option value="${defaultUnit}" ${currentUnit === defaultUnit ? 'selected' : ''}>${defaultUnit}</option>
                    <option value="${altUnit}" ${currentUnit === altUnit ? 'selected' : ''}>${altUnit}</option>
                `;
            }
        });
    }

    updatePriceForSelectedUnit(resourceName, selectedUnit) {
        const resource = this.getResourceInfo(resourceName);
        if (!resource) return;

        const defaultUnit = resource.Unit;
        // Handle NaN values properly
        const customPrice = this.customPrices.get(resourceName);
        const unitCost = resource['Unit Cost'];
        const defaultPrice = customPrice || ((unitCost && !isNaN(unitCost)) ? unitCost : 0);
        
        // Find all price inputs for this resource
        const priceInputs = document.querySelectorAll(`.price-input[data-resource="${resourceName}"]`);
        
        priceInputs.forEach(input => {
            let displayPrice;
            
            if (selectedUnit === defaultUnit) {
                // Show default unit price
                displayPrice = defaultPrice;
            } else {
                // Show alternative unit price
                displayPrice = this.calculateAltUnitPrice(resourceName, defaultPrice, defaultUnit, selectedUnit);
                

            }
            
            // Update the input value - ensure we don't get NaN
            const safeDisplayPrice = isNaN(displayPrice) ? 0 : displayPrice;
            input.value = safeDisplayPrice;

            // If labor with extra-per-floor, re-pin base so displayed price remains fixed for current floor
            const isLabor = input.dataset.type === 'labor';
            if (isLabor && this.isLaborWithFloorExtra(resourceName)) {
                const row = input.closest('tr');
                const extraEl = row ? row.querySelector('.extra-per-floor-input') : null;
                const floorLevel = parseInt(this.laborFloorLevelInput ? this.laborFloorLevelInput.value : '1') || 1;
                const extra = extraEl ? (parseFloat(extraEl.value) || 0) : 0;
                const computedBaseForFloor1 = safeDisplayPrice - extra * (floorLevel - 1);
                input.dataset.base = isNaN(computedBaseForFloor1) ? '0' : String(computedBaseForFloor1);
            }
        });
    }

    setupEventListeners() {
        this.mainItemSelect.addEventListener('change', () => {
            const selectedMainItem = this.mainItemSelect.value;
            if (selectedMainItem) {
                this.loadSubItems(selectedMainItem);
            } else {
                this.subItemSelect.innerHTML = '<option value="">-- اختر البند الفرعي --</option>';
                this.subItemSelect.disabled = true;
            }
            this.calculate();
        });

        this.subItemSelect.addEventListener('change', () => {
            this.calculate();
        });

        this.quantityInput.addEventListener('focus', (e) => {
            e.target.select();
        });
        
        this.quantityInput.addEventListener('input', () => {
            this.calculate();
        });

        // Waste percent input event
        if (this.wastePercentInput) {
            this.wastePercentInput.addEventListener('input', () => {
                this.calculate();
                // Also update existing summary cards to reflect new waste percentage
                this.updateAllSummaryCardsWithNewPercentages();
                // Save the global percentage to project
                this.saveGlobalPercentages();
            });
        }

        // Operation percent input event
        if (this.operationPercentInput) {
            this.operationPercentInput.addEventListener('input', () => {
                this.calculate();
                // Also update existing summary cards to reflect new operation percentage
                this.updateAllSummaryCardsWithNewPercentages();
                // Save the global percentage to project
                this.saveGlobalPercentages();
            });
        }

        // Setup tab functionality
        this.setupTabs();

        // Accordion logic
        const accordionSections = [
            { header: this.materialsHeader, body: this.materialsBody },
            { header: this.workmanshipHeader, body: this.workmanshipBody },
            { header: this.laborHeader, body: this.laborBody },
        ];
        accordionSections.forEach(({ header, body }) => {
            if (!header || !body) return;
            const toggleBtn = header.querySelector('.accordion-toggle');
            const toggle = () => {
                const expanded = toggleBtn.getAttribute('aria-expanded') === 'true';
                if (expanded) {
                    body.style.display = 'none';
                    toggleBtn.setAttribute('aria-expanded', 'false');
                    toggleBtn.textContent = '+';
                } else {
                    body.style.display = 'block';
                    toggleBtn.setAttribute('aria-expanded', 'true');
                    toggleBtn.textContent = '−';
                }
            };
            // Collapse by default
            body.style.display = 'none';
            toggleBtn.setAttribute('aria-expanded', 'false');
            toggleBtn.textContent = '+';
            header.addEventListener('click', (e) => {
                // Only toggle if not clicking inside a button/input
                if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'BUTTON' && e.target.tagName !== 'SELECT') {
                    toggle();
                }
            });
            toggleBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                toggle();
            });
        });

        if (this.saveItemBtn) {
            this.saveItemBtn.addEventListener('click', () => this.saveCurrentItemToSummary());
        }
        if (this.summaryUndoBtn) {
            this.summaryUndoBtn.addEventListener('click', () => this.undoLastDelete());
        }

        // Update quantity input placeholder/unit live
        const updateQuantityUnit = () => {
            const mainItem = this.mainItemSelect.value;
            const subItem = this.subItemSelect.value;
            const unit = this.getCorrectUnit(mainItem, subItem);
            if (this.quantityInput) {
                this.quantityInput.placeholder = unit ? `أدخل الكمية (${unit})` : 'أدخل الكمية';
            }
        };
        if (this.mainItemSelect) this.mainItemSelect.addEventListener('change', updateQuantityUnit);
        if (this.subItemSelect) this.subItemSelect.addEventListener('change', updateQuantityUnit);

        if (this.resourceTypeFilter) {
            this.resourceTypeFilter.addEventListener('change', () => this.renderResourcesSummary());
        }
    }

    setupTabs() {
        const tabButtons = document.querySelectorAll('.tab-btn');
        const tabContents = document.querySelectorAll('.tab-content');

        tabButtons.forEach(button => {
            button.addEventListener('click', () => {
                const targetTab = button.dataset.tab;
                
                // Remove active class from all buttons and contents
                tabButtons.forEach(btn => btn.classList.remove('active'));
                tabContents.forEach(content => content.classList.remove('active'));
                
                // Add active class to clicked button and corresponding content
                button.classList.add('active');
                document.getElementById(`${targetTab}-tab`).classList.add('active');
            });
        });
    }

    getResourceInfo(resourceName, preferredType = null) {
        // If we have a preferred type, try to find the resource in that specific list first
        if (preferredType) {
            let targetList;
            if (preferredType === 'خامات') {
                targetList = materialsList;
            } else if (preferredType === 'مصنعيات') {
                targetList = workmanshipList;
            } else if (preferredType === 'عمالة') {
                targetList = laborList;
            }
            
            if (targetList) {
                const resource = targetList.find(resource => resource.Resource === resourceName);
                if (resource) {
                    return resource;
                }
            }
        }
        
        // Fallback to the combined list
        return this.resourcesList.find(resource => resource.Resource === resourceName);
    }

    getResourcePrice(resourceName, preferredType = null) {
        const resourceInfo = this.getResourceInfo(resourceName, preferredType);
        if (!resourceInfo) return 0;

        // Always use default unit for calculations, regardless of display unit
        // This ensures consistent calculations even when alternative units are displayed
        const defaultUnit = resourceInfo.Unit;

        // Get the price in the default unit
        let price;
        if (this.customPrices.has(resourceName)) {
            // Custom prices are always stored in default unit
            price = this.customPrices.get(resourceName);
        } else {
            // Handle NaN values properly
            const unitCost = resourceInfo['Unit Cost'];
            price = (unitCost && !isNaN(unitCost)) ? unitCost : 0;
        }

        return price;
    }

    getResourceUnit(resourceName, preferredType = null) {
        const resourceInfo = this.getResourceInfo(resourceName, preferredType);
        if (!resourceInfo) return '';

        // For تفاصيل التكلفة, always return the base unit (ignore custom units)
        // This ensures تفاصيل التكلفة always shows base units like لتر, شيكارة, م3
        return resourceInfo.Unit;
    }

    getResourceDisplayUnit(resourceName, preferredType = null) {
        const resourceInfo = this.getResourceInfo(resourceName, preferredType);
        if (!resourceInfo) return '';

        // For أسعار الموارد, return custom unit if set, otherwise base unit
        // This allows أسعار الموارد to show both units freely
        return this.customUnits.get(resourceName) || resourceInfo.Unit;
    }

    calculate() {
        const mainItem = this.mainItemSelect.value;
        const subItem = this.subItemSelect.value;
        const quantity = parseFloat(this.quantityInput.value) || 0;
        const wastePercent = this.wastePercentInput ? (parseFloat(this.wastePercentInput.value) || 0) : 0;
        const operationPercent = this.operationPercentInput ? (parseFloat(this.operationPercentInput.value) || 0) : 0;

        if (!mainItem || !subItem || quantity <= 0) {
            this.totalCostElement.textContent = '0.00';
            this.resultsSection.style.display = 'none';
            // Update resources totals even when no calculation
            this.updateResourcesTotals();
            return;
        }

        // Get all items that match the selected main and sub item
        const matchingItems = itemsList.filter(item => 
            item['Main Item'] === mainItem && 
            item['Sub Item'] === subItem
        );

        if (matchingItems.length === 0) {
            this.totalCostElement.textContent = '0.00';
            this.resultsSection.style.display = 'none';
            // Update resources totals even when no matching items
            this.updateResourcesTotals();
            return;
        }

        // Group items by type (خامات, مصنعيات, عمالة)
        const materials = [];
        const workmanship = [];
        const labor = [];

        matchingItems.forEach(item => {
            const resourceInfo = this.getResourceInfo(item.Resource, item.Type);
            if (resourceInfo) {
                const unitPrice = this.getResourcePrice(item.Resource, item.Type);
                // Use custom rate if set, otherwise default
                const rateKey = `${mainItem}||${subItem}||${item.Resource}`;
                const itemQuantity = this.customRates[rateKey] !== undefined ? parseFloat(this.customRates[rateKey]) : (parseFloat(item['Quantity per Unit']) || 0);
                const totalQuantity = itemQuantity * quantity;
                // Ensure we don't get NaN in calculations
                const safeUnitPrice = isNaN(unitPrice) ? 0 : unitPrice;
                const totalCost = safeUnitPrice * totalQuantity;
                // Always use default unit for calculations, but show custom unit for display
                const defaultUnit = resourceInfo.Unit;
                const displayUnit = this.getResourceUnit(item.Resource, item.Type);

                const itemData = {
                    resource: item.Resource,
                    quantity: totalQuantity,
                    unit: displayUnit, // Show the selected unit for display
                    unitPrice: safeUnitPrice,
                    totalCost: totalCost,
                    type: resourceInfo.Type,
                    rateKey,
                    rate: itemQuantity,
                    defaultRate: parseFloat(item['Quantity per Unit']) || 0
                };

                if (resourceInfo.Type === 'خامات') {
                    materials.push(itemData);
                } else if (resourceInfo.Type === 'مصنعيات') {
                    workmanship.push(itemData);
                } else if (resourceInfo.Type === 'عمالة') {
                    labor.push(itemData);
                }
            }
        });

        // Calculate totals
        const materialsTotal = materials.reduce((sum, item) => sum + item.totalCost, 0);
        const workmanshipTotal = workmanship.reduce((sum, item) => sum + item.totalCost, 0);
        const laborTotal = labor.reduce((sum, item) => sum + item.totalCost, 0);
        let baseTotal = materialsTotal + workmanshipTotal + laborTotal;

        // Apply waste percent to base, then add operation percent on base only
        let grandTotal = (baseTotal * (1 + wastePercent / 100)) + (baseTotal * (operationPercent / 100));

        // Update display with formatted numbers
        this.totalCostElement.textContent = this.formatNumber(grandTotal);
        this.resultsSection.style.display = 'block';

        // Update tables
        this.updateTable(this.materialsTable, materials, 'خامات');
        this.updateTable(this.workmanshipTable, workmanship, 'مصنعيات');
        this.updateTable(this.laborTable, labor, 'عمالة');

        // Update accordion headers with description and totals
        if (this.materialsDesc && this.materialsTotal) {
            this.materialsDesc.textContent = `عدد البنود: ${materials.length}`;
            this.materialsTotal.textContent = `${this.formatNumber(materialsTotal)} جنيه`;
        }
        if (this.workmanshipDesc && this.workmanshipTotal) {
            this.workmanshipDesc.textContent = `عدد البنود: ${workmanship.length}`;
            this.workmanshipTotal.textContent = `${this.formatNumber(workmanshipTotal)} جنيه`;
        }
        if (this.laborDesc && this.laborTotal) {
            this.laborDesc.textContent = `عدد البنود: ${labor.length}`;
            this.laborTotal.textContent = `${this.formatNumber(laborTotal)} جنيه`;
        }

        // Calculate unit price for display
        let unitPrice = 0;
        let unit = '';
        if (quantity > 0 && grandTotal > 0) {
            unitPrice = grandTotal / quantity;
            unit = this.getCorrectUnit(mainItem, subItem);
        }
        if (this.unitPriceDisplay) {
            if (unitPrice > 0 && unit) {
                this.unitPriceDisplay.innerHTML = `<span class="label">تكلفة الوحدة:</span> <span class="value">${this.formatNumber(unitPrice)} جنيه / ${unit}</span>`;
            } else {
                this.unitPriceDisplay.textContent = '';
            }
        }

        // Update summary total
        this.updateSummaryTotal();
        
        // Update resources totals - ALWAYS call this
        this.updateResourcesTotals();
        
        // Store the calculated unit price for use in summary cards
        this.lastCalculatedUnitPrice = unitPrice;
        this.lastCalculatedUnit = unit;
    }

    updateTable(tableBody, items, type) {
        tableBody.innerHTML = '';

        // Add table header
        const thead = tableBody.parentElement.querySelector('thead');
        if (thead) {
            thead.innerHTML = `<tr>
                <th>الخامة</th>
                <th>معدل الاستخدام</th>
                <th>الكمية المطلوبة</th>
                <th>الوحدة</th>
                <th>سعر الوحدة</th>
                <th>التكلفة</th>
            </tr>`;
        }

        if (items.length === 0) {
            const row = tableBody.insertRow();
            const cell = row.insertCell();
            cell.colSpan = 6;
            cell.textContent = 'لا توجد عناصر';
            cell.style.textAlign = 'center';
            cell.style.color = '#666';
            return;
        }

        let total = 0;
        items.forEach(item => {
            const row = tableBody.insertRow();
            const resourceCell = row.insertCell();
            resourceCell.textContent = item.resource;
            // Editable rate input
            const rateCell = row.insertCell();
            rateCell.style.textAlign = 'center';
            rateCell.innerHTML = `
                <input type="number" class="rate-input" value="${this.formatRate(item.rate)}" min="0" step="0.0001" style="width:70px;"> 
                <button class="rate-default-btn" title="الافتراضي" style="margin-right:4px;">↺</button>
            `;
            // Add event listeners
            const rateInput = rateCell.querySelector('.rate-input');
            const defaultBtn = rateCell.querySelector('.rate-default-btn');
            rateInput.addEventListener('input', (e) => {
                this.customRates[item.rateKey] = e.target.value;
                this.saveProjectCustomRates();
                this.calculate();
                // Update resources totals ribbon live
                this.updateResourcesTotals();
            });
            defaultBtn.addEventListener('click', (e) => {
                e.preventDefault();
                this.customRates[item.rateKey] = this.formatRate(item.defaultRate);
                this.saveProjectCustomRates();
                this.calculate();
                // Update resources totals ribbon live
                this.updateResourcesTotals();
            });
            // الكمية المطلوبة
            const quantityCell = row.insertCell();
            quantityCell.textContent = this.formatNumber(item.quantity);
            // الوحدة
            const unitCell = row.insertCell();
            unitCell.textContent = item.unit;
            // سعر الوحدة
            const unitPriceCell = row.insertCell();
            unitPriceCell.textContent = `${this.formatNumber(item.unitPrice)} جنيه`;
            // التكلفة
            const totalCostCell = row.insertCell();
            totalCostCell.textContent = `${this.formatNumber(item.totalCost)} جنيه`;
            total += isNaN(item.totalCost) ? 0 : item.totalCost;
        });

        // Add total row at the bottom
        const totalRow = tableBody.insertRow();
        totalRow.style.fontWeight = 'bold';
        const labelCell = totalRow.insertCell();
        labelCell.colSpan = 5;
        if (type === 'خامات') {
            labelCell.textContent = 'إجمالي الخامات';
        } else if (type === 'مصنعيات') {
            labelCell.textContent = 'إجمالي المصنعيات';
        } else if (type === 'عمالة') {
            labelCell.textContent = 'إجمالي العمالة';
        } else {
            labelCell.textContent = 'الإجمالي';
        }
        const totalValueCell = totalRow.insertCell();
        totalValueCell.textContent = `${this.formatNumber(total)} جنيه`;
    }

    // Helper to check if resource needs extra per floor
    isLaborWithFloorExtra(resourceName) {
        return [
            'تشوين أسمنت',
            'تشوين رمل',
            'تشوين طوب',
            'تشوين مادة لاصقة',
        ].includes(resourceName);
    }

    // Update prices for labor items when floor level changes (labor-only)
    updateLaborPricesForFloor(previousFloorLevel) {
        try {
        const currentFloor = (this.laborFloorLevelInput && this.laborFloorLevelInput.value)
            ? (parseInt(this.laborFloorLevelInput.value) || 1)
            : (this.laborFloorLevel || 1);
        const prevFloor = previousFloorLevel || currentFloor;
            
            // Update this.laborFloorLevel to keep it in sync
            this.laborFloorLevel = currentFloor;
            
        // 1) Update visible inputs if present
        Object.entries(this.laborExtraInputs).forEach(([resourceName, extraInput]) => {
            const priceInput = document.querySelector(`.price-input[data-resource='${resourceName}'][data-type='labor']`);
            let extra = parseFloat(extraInput.value);
            if (isNaN(extra)) {
                extra = (this.laborExtrasPerFloor && this.laborExtrasPerFloor.hasOwnProperty(resourceName)) ? (parseFloat(this.laborExtrasPerFloor[resourceName]) || 0) : 0;
                extraInput.value = String(extra);
            }
            if (priceInput) {
                const baseFromDataset = parseFloat(priceInput.dataset.base);
                const base = !isNaN(baseFromDataset) ? baseFromDataset : ((parseFloat(priceInput.value) || 0) - extra * (prevFloor - 1));
                const newPrice = base + extra * (currentFloor - 1);
                    priceInput.value = isNaN(newPrice) ? 0 : newPrice;
                    this.setCustomPrice(resourceName, parseFloat(priceInput.value.replace(/,/g, '')) || 0);
            }
        });
            
        // 2) Update stored prices even if inputs are not mounted (using previous stored price as baseline)
        Object.keys(this.laborExtrasPerFloor || {}).forEach(resourceName => {
            const extra = parseFloat(this.laborExtrasPerFloor[resourceName]) || 0;
            // Get currently stored default-unit price for this resource
            const storedPrice = this.getResourcePrice(resourceName, 'عمالة');
            if (storedPrice !== undefined && storedPrice !== null) {
                const base = storedPrice - extra * (prevFloor - 1);
                const newPrice = base + extra * (currentFloor - 1);
                this.setCustomPrice(resourceName, isNaN(newPrice) ? 0 : newPrice);
            }
        });
            
            // Save the current floor level to project data
            this.saveProjectLaborFloorLevel();
            
        this.calculate();
            this.updateResourcesTotals();
            
        } catch (error) {
            console.error('Error updating labor prices for floor:', error);
        }
    }

    saveCurrentItemToSummary() {
        // Get current selection and calculation
        const mainItem = this.mainItemSelect.value;
        const subItem = this.subItemSelect.value;
        const quantity = parseFloat(this.quantityInput.value) || 0;
        const total = parseFloat(this.totalCostElement.textContent.replace(/,/g, '')) || 0;
        if (!mainItem || !subItem || quantity <= 0 || total <= 0) return;
        
        // Use the stored unit price from the calculation for consistency
        const unitPrice = this.lastCalculatedUnitPrice || (total / quantity);
        const unit = this.lastCalculatedUnit || this.getCorrectUnit(mainItem, subItem);
        
        // Get current waste and operation percentages from input fields
        const wastePercent = this.wastePercentInput ? (parseFloat(this.wastePercentInput.value) || 0) : 0;
        const operationPercent = this.operationPercentInput ? (parseFloat(this.operationPercentInput.value) || 0) : 0;
        
        // Create card data
        const cardData = {
            id: Date.now() + Math.random(),
            mainItem,
            subItem,
            quantity,
            unit,
            total,
            unitPrice,
            wastePercent,
            operationPercent
        };
        this.addSummaryCard(cardData);
    }

    addSummaryCard(cardData) {
        // Use saved values or defaults
        const initialTax = cardData.taxPercentage !== undefined ? cardData.taxPercentage : 14;
        const initialRisk = cardData.riskPercentage !== undefined ? cardData.riskPercentage : 0;
        
        // If there's a saved sell price, use it; otherwise calculate it
        let initialSellPrice;
        if (cardData.sellPrice !== undefined && cardData.sellPrice !== null) {
            initialSellPrice = cardData.sellPrice;
        } else {
            initialSellPrice = cardData.unitPrice * (1 + initialRisk / 100) * (1 + initialTax / 100);
        }
        
        // Create card element
        const card = document.createElement('div');
        card.className = 'summary-card';
        card.dataset.cardId = cardData.id;
        card.cardData = cardData;
        
        // Set default selection state (new items are selected by default)
        if (cardData.selected === undefined) {
            cardData.selected = true;
        }
        // Collapsible content
        card.innerHTML = `
            <div class="summary-card-header">
                <input type="checkbox" class="item-checkbox" ${cardData.selected ? 'checked' : ''}>
                <span class="expand-icon">&#9654;</span>
                <div class="item-title">${cardData.mainItem} - ${cardData.subItem}</div>
                <div class="item-details">الكمية: <input type="number" class="quantity-input" value="${cardData.quantity}" min="0" step="0.01" style="width: 80px; text-align: center;"> ${cardData.unit}</div>
                <div class="item-unit">تكلفة الوحدة: <span class="unit-price-value">${this.formatNumber(cardData.unitPrice)}</span> جنيه / ${cardData.unit}</div>
                <div class="item-sell-header">سعر البيع: <span class="sell-price-header">${this.formatNumber(initialSellPrice)}</span> جنيه / ${cardData.unit}</div>
                <div class="item-total">الإجمالي: ${this.formatNumber(initialSellPrice * cardData.quantity)} جنيه</div>
                <button class="delete-btn">حذف</button>
            </div>
            <div class="summary-card-body">
                <div class="card-row">
                    <div class="input-group">
                        <label>نسبة المخاطر (%)</label>
                        <input type="number" class="risk-input" min="0" step="0.01" value="${initialRisk}">
                    </div>
                    <div class="input-group">
                        <label>نسبة الضريبة (%)</label>
                        <input type="number" class="tax-input" min="0" step="0.01" value="${initialTax}">
                    </div>
                    <div class="input-group">
                        <label>نسبة الهالك (%)</label>
                        <input type="number" class="waste-input" min="0" step="0.01" value="${cardData.wastePercent || 0}">
                    </div>
                    <div class="input-group">
                        <label>نسبة التشغيل (%)</label>
                        <input type="number" class="operation-input" min="0" step="0.01" value="${cardData.operationPercent || 0}">
                    </div>
                    <div class="display-group">
                        <span class="item-unit">تكلفة الوحدة: <span class="unit-price-value">${this.formatNumber(cardData.unitPrice)}</span> جنيه / ${cardData.unit}</span>
                    </div>
                    <div class="display-group">
                        <span class="item-sell">سعر البيع: <span class="sell-price-value">${this.formatNumber(initialSellPrice)}</span> جنيه / ${cardData.unit}</span>
                    </div>
                    <div class="display-group">
                        <span class="item-total-body">الإجمالي: <span class="body-total-value">${this.formatNumber(initialSellPrice * cardData.quantity)}</span> جنيه</span>
                    </div>
                </div>
            </div>
        `;
        // Expand/collapse logic
        const header = card.querySelector('.summary-card-header');
        const body = card.querySelector('.summary-card-body');
        const expandIcon = card.querySelector('.expand-icon');
        let expanded = false;
        const setExpanded = (val) => {
            expanded = val;
            if (expanded) {
                body.style.display = 'block';
                expandIcon.innerHTML = '&#9660;';
            } else {
                body.style.display = 'none';
                expandIcon.innerHTML = '&#9654;';
            }
        };
        setExpanded(false);
        header.addEventListener('click', (e) => {
            // Only toggle if not clicking delete
            if (!e.target.classList.contains('delete-btn')) setExpanded(!expanded);
        });
        // Delete logic
        const deleteBtn = card.querySelector('.delete-btn');
        deleteBtn.addEventListener('click', () => {
            // Add to undo stack
            this.addToUndoStack({
                action: 'delete',
                cardData: cardData,
                cardElement: card,
                position: this.getCardPosition(card)
            });
            
            // Remove the card
            card.remove();
            
            // Update project items after delete FIRST
            this.saveProjectItemsFromDOM();
            
            // Update selection count after deleting card
            this.updateSelectionCount();
            
            // Then update all totals and displays
            this.updateSummaryTotal();
            this.updateSummarySellingTotal();
            this.updateSummaryFinalTotal();
            this.updateResourcesSection();
            
            // Update resources totals ribbon after all data is saved
            this.updateResourcesTotals();
            
            // Show undo button
            if (this.summaryUndoBtn) this.summaryUndoBtn.style.display = 'inline-block';
        });
        // Risk/Tax logic
        const riskInput = card.querySelector('.risk-input');
        const taxInput = card.querySelector('.tax-input');
        const unitPriceValue = card.querySelector('.unit-price-value');
        const sellPriceValue = card.querySelector('.sell-price-value');
        const sellPriceHeader = card.querySelector('.sell-price-header');
        const updateSellPrice = () => {
            const risk = parseFloat(riskInput.value) || 0;
            const tax = parseFloat(taxInput.value) || 0;
            const base = cardData.unitPrice; // Use the stored unit price
            const sell = base * (1 + risk / 100) * (1 + tax / 100);
            sellPriceValue.textContent = this.formatNumber(sell);
            sellPriceHeader.textContent = this.formatNumber(sell);
            card.dataset.sellPrice = sell;
            
            // Save risk and tax values to card data
            cardData.riskPercentage = risk;
            cardData.taxPercentage = tax;
            cardData.sellPrice = sell;
            
            // Update the card's total display live
            this.updateCardTotalDisplay(card, cardData);
            
            // Save project data to persist the changes FIRST
            this.saveProjectItemsFromDOM();
            
            // Then update all totals
            this.updateSummaryTotal();
            this.updateSummarySellingTotal();
            this.updateSummaryFinalTotal();
            
            // Update resources section and totals
            this.updateResourcesSection();
            this.updateResourcesTotals();
        };
        riskInput.addEventListener('input', updateSellPrice);
        taxInput.addEventListener('input', updateSellPrice);
        
        // Waste and Operation percentage event listeners
        const wasteInput = card.querySelector('.waste-input');
        const operationInput = card.querySelector('.operation-input');
        
        wasteInput.addEventListener('input', (e) => {
            const wastePercent = parseFloat(e.target.value) || 0;
            cardData.wastePercent = wastePercent;
            
            // Update the card's total display live
            this.updateCardTotalDisplay(card, cardData);
            
            this.saveProjectItemsFromDOM();
        this.updateSummaryTotal();
            this.updateSummarySellingTotal();
            this.updateSummaryFinalTotal();
        this.updateResourcesSection();
            this.updateResourcesTotals();
        });
        
        // Select all text when clicking on waste input
        wasteInput.addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent card from expanding/collapsing
            e.target.select();
        });
        
        operationInput.addEventListener('input', (e) => {
            const operationPercent = parseFloat(e.target.value) || 0;
            cardData.operationPercent = operationPercent;
            
            // Update the card's total display live
            this.updateCardTotalDisplay(card, cardData);
            
        this.saveProjectItemsFromDOM();
            this.updateSummaryTotal();
            this.updateSummarySellingTotal();
            this.updateSummaryFinalTotal();
            this.updateResourcesSection();
            this.updateResourcesTotals();
        });
        
        // Select all text when clicking on operation input
        operationInput.addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent card from expanding/collapsing
            e.target.select();
        });
        
        // Individual card checkbox event listener
        const itemCheckbox = card.querySelector('.item-checkbox');
        itemCheckbox.addEventListener('change', (e) => {
            e.stopPropagation(); // Prevent card expansion
            cardData.selected = e.target.checked;
            this.updateSelectionCount();
            this.saveProjectItemsFromDOM();
            
            // Update all totals immediately after individual selection change
            this.updateSummaryTotal();
            this.updateSummarySellingTotal();
            this.updateSummaryFinalTotal();
            this.updateResourcesSection();
        });
        
        // Prevent card expansion when clicking checkbox
        itemCheckbox.addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent card from expanding/collapsing
        });
        
        // Quantity input event listener for live calculations
        const quantityInput = card.querySelector('.quantity-input');
        
        // Select all text when clicking on input
        quantityInput.addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent card from expanding/collapsing
            e.target.select();
        });
        
        // Limit to 2 decimal places and handle live calculations
        quantityInput.addEventListener('input', (e) => {
            let value = e.target.value;
            
            // Limit to 2 decimal places
            if (value.includes('.')) {
                const parts = value.split('.');
                if (parts[1] && parts[1].length > 2) {
                    value = parseFloat(value).toFixed(2);
                    e.target.value = value;
                }
            }
            
            const newQuantity = parseFloat(value) || 0;
            
            // Update card data
            cardData.quantity = newQuantity;
            
            // Update the card's total display live
            this.updateCardTotalDisplay(card, cardData);
            
            // Save project data
            this.saveProjectItemsFromDOM();
            
            // Update all summary totals
            this.updateSummaryTotal();
            this.updateSummarySellingTotal();
            this.updateSummaryFinalTotal();
            
            // Update resources section and totals
            this.updateResourcesSection();
            this.updateResourcesTotals();
        });
        
        // The inputs already have the correct initial values from the HTML template
        // Just ensure the card data is updated with the initial values
        if (cardData.riskPercentage === undefined) {
            cardData.riskPercentage = initialRisk;
        }
        if (cardData.taxPercentage === undefined) {
            cardData.taxPercentage = initialTax;
        }
        
        // The sell price is already calculated and displayed correctly
        // Just ensure the card dataset is updated
        card.dataset.sellPrice = initialSellPrice;
        
        // Update the card's total display with initial values
        this.updateCardTotalDisplay(card, cardData);
        
        this.summaryCards.appendChild(card);
        
        // Update selection count after adding new card
        this.updateSelectionCount();
        
        // Save all summary cards to project after add FIRST
        this.saveProjectItemsFromDOM();
        
        // Then update all totals
        this.updateSummaryTotal();
        this.updateSummaryFinalTotal();
        
        // Update resources totals ribbon after data is saved
        this.updateResourcesTotals();

        // Add 'عرض التفاصيل' button to summary card
        const detailsBtn = document.createElement('button');
        detailsBtn.textContent = 'عرض التفاصيل';
        detailsBtn.className = 'details-btn';
        detailsBtn.style.marginRight = '12px';
        detailsBtn.onclick = (e) => {
            e.stopPropagation();
            this.showItemDetailsModal(cardData);
        };
        header.insertBefore(detailsBtn, header.firstChild);
    }

    // Global function to enhance number inputs
    enhanceNumberInputs() {
        // Select all text when clicking on any number input
        document.addEventListener('click', (e) => {
            if (e.target.type === 'number') {
                e.target.select();
            }
        });
        
        // Limit all number inputs to 2 decimal places
        document.addEventListener('input', (e) => {
            if (e.target.type === 'number') {
                let value = e.target.value;
                if (value.includes('.')) {
                    const parts = value.split('.');
                    if (parts[1] && parts[1].length > 2) {
                        value = parseFloat(value).toFixed(2);
                        e.target.value = value;
                    }
                }
            }
        });
    }

    // Ensure all cards have proper sellPrice in dataset
    ensureAllCardsHaveSellPrice() {
        const cards = this.summaryCards.querySelectorAll('.summary-card');
        cards.forEach(card => {
            if (!card.dataset.sellPrice || card.dataset.sellPrice === 'undefined') {
                const cardData = card.cardData;
                if (cardData) {
                    const taxPercentage = cardData.taxPercentage !== undefined ? cardData.taxPercentage : 14;
                    const riskPercentage = cardData.riskPercentage !== undefined ? cardData.riskPercentage : 0;
                    const sellPrice = cardData.unitPrice * (1 + riskPercentage / 100) * (1 + taxPercentage / 100);
                    card.dataset.sellPrice = sellPrice;
                }
            }
        });
    }

    // Setup selection system for summary cards
    setupSelectionSystem() {
        console.log('Setting up selection system...');
        
        // Wait a bit for DOM to be ready
        setTimeout(() => {
            const selectAllCheckbox = document.getElementById('selectAllCheckbox');
            const selectionCount = document.getElementById('selectionCount');
            const totalItems = document.getElementById('totalItems');
            
            console.log('Selection elements found:', { selectAllCheckbox, selectionCount, totalItems });
            
            if (!selectAllCheckbox || !selectionCount || !totalItems) {
                console.warn('Some selection elements not found, retrying...');
                setTimeout(() => this.setupSelectionSystem(), 100);
                return;
            }
            
            // Remove any existing event listeners
            const newSelectAllCheckbox = selectAllCheckbox.cloneNode(true);
            selectAllCheckbox.parentNode.replaceChild(newSelectAllCheckbox, selectAllCheckbox);
            
                    // Select all functionality
        newSelectAllCheckbox.addEventListener('change', (e) => {
            console.log('Select all checkbox changed:', e.target.checked);
            const isChecked = e.target.checked;
            const cards = this.summaryCards.querySelectorAll('.summary-card');
            
            console.log('Found cards to update:', cards.length);
            
            cards.forEach((card, index) => {
                const checkbox = card.querySelector('.item-checkbox');
                if (checkbox) {
                    checkbox.checked = isChecked;
                    if (card.cardData) {
                        card.cardData.selected = isChecked;
                    }
                    console.log(`Updated card ${index + 1}:`, isChecked);
                }
            });
            
            this.updateSelectionCount();
            this.saveProjectItemsFromDOM();
            
            // Update all totals immediately after selection change
            this.updateSummaryTotal();
            this.updateSummarySellingTotal();
            this.updateSummaryFinalTotal();
            this.updateResourcesSection();
        });
            
            // Update selection count
            this.updateSelectionCount();
            console.log('Selection system setup complete');
        }, 200);
    }

    // Update selection count display
    updateSelectionCount() {
        const selectionCount = document.getElementById('selectionCount');
        const totalItems = document.getElementById('totalItems');
        const selectAllCheckbox = document.getElementById('selectAllCheckbox');
        
        if (!selectionCount || !totalItems || !selectAllCheckbox) {
            console.warn('Selection elements not found in updateSelectionCount');
            return;
        }
        
        const cards = this.summaryCards.querySelectorAll('.summary-card');
        const selectedCards = this.summaryCards.querySelectorAll('.summary-card .item-checkbox:checked');
        
        console.log('Updating selection count:', { cards: cards.length, selected: selectedCards.length });
        
        totalItems.textContent = cards.length;
        selectionCount.textContent = selectedCards.length;
        
        // Update select all checkbox state
        if (cards.length === 0) {
            selectAllCheckbox.checked = false;
            selectAllCheckbox.indeterminate = false;
        } else if (selectedCards.length === 0) {
            selectAllCheckbox.checked = false;
            selectAllCheckbox.indeterminate = false;
        } else if (selectedCards.length === cards.length) {
            selectAllCheckbox.checked = true;
            selectAllCheckbox.indeterminate = false;
        } else {
            selectAllCheckbox.checked = false;
            selectAllCheckbox.indeterminate = true;
        }
        
        console.log('Select all checkbox state updated:', { 
            checked: selectAllCheckbox.checked, 
            indeterminate: selectAllCheckbox.indeterminate 
        });
        
        // Update resources totals after selection count changes
        this.updateResourcesTotals();
    }

    // Get selected cards for export and resources management
    getSelectedCards() {
        const selectedCards = this.summaryCards.querySelectorAll('.summary-card .item-checkbox:checked');
        return Array.from(selectedCards).map(checkbox => checkbox.closest('.summary-card'));
    }

    // Save global waste and operation percentages to project
    saveGlobalPercentages() {
        if (!this.currentProjectId || !this.projects[this.currentProjectId]) return;
        
        const proj = this.projects[this.currentProjectId];
        proj.globalWastePercent = parseFloat(this.wastePercentInput?.value) || 0;
        proj.globalOperationPercent = parseFloat(this.operationPercentInput?.value) || 0;
        
        this.saveProjects();
        console.log('Saved global percentages to project:', { 
            waste: proj.globalWastePercent, 
            operation: proj.globalOperationPercent 
        });
    }

    // Update all summary cards with new global waste/operation percentages
    updateAllSummaryCardsWithNewPercentages() {
        try {
            console.log('=== updateAllSummaryCardsWithNewPercentages called ===');
            
            const cards = this.summaryCards.querySelectorAll('.summary-card');
            const globalWastePercent = parseFloat(this.wastePercentInput?.value) || 0;
            const globalOperationPercent = parseFloat(this.operationPercentInput?.value) || 0;
            
            console.log('Global percentages:', { globalWastePercent, globalOperationPercent });
            console.log('Found cards to update:', cards.length);
            
            cards.forEach((card, index) => {
                const cardData = card.cardData;
                if (cardData) {
                    console.log(`Updating card ${index + 1}:`, cardData.mainItem, cardData.subItem);
                    
                    // Update card data with new global percentages
                    cardData.wastePercent = globalWastePercent;
                    cardData.operationPercent = globalOperationPercent;
                    
                    // Update the card's display
                    this.updateCardTotalDisplay(card, cardData);
                }
            });
            
            // Update all summary totals
            this.updateSummaryTotal();
            this.updateSummarySellingTotal();
            this.updateSummaryFinalTotal();
            this.updateResourcesSection();
            this.updateResourcesTotals();
            
            // Save the updated data
            this.saveProjectItemsFromDOM();
            
            console.log('=== All cards updated successfully ===');
            
        } catch (error) {
            console.error('Error updating summary cards with new percentages:', error);
        }
    }

    // Update individual card's total display with proper calculations
    updateCardTotalDisplay(card, cardData) {
        try {
            console.log('=== updateCardTotalDisplay called ===');
            console.log('Card data:', cardData);
            
            const unitPrice = parseFloat(cardData.unitPrice) || 0;
            const quantity = parseFloat(cardData.quantity) || 0;
            const wastePercent = parseFloat(cardData.wastePercent) || 0;
            const operationPercent = parseFloat(cardData.operationPercent) || 0;
            const riskPercentage = parseFloat(cardData.riskPercentage) || 0;
            const taxPercentage = parseFloat(cardData.taxPercentage) || 14;
            
            // Calculate adjusted unit cost with waste and operation (same as ادخال الكميات)
            const wasteAmount = unitPrice * (wastePercent / 100);
            const operationAmount = unitPrice * (operationPercent / 100);
            const adjustedUnitCost = unitPrice + wasteAmount + operationAmount;
            
            console.log('Calculations:', { unitPrice, wastePercent, operationPercent, wasteAmount, operationAmount, adjustedUnitCost });
            
            // Calculate selling price with risk and tax
            const sellPrice = unitPrice * (1 + riskPercentage / 100) * (1 + taxPercentage / 100);
            
            // Calculate total with adjusted unit cost
            const cardTotal = adjustedUnitCost * quantity;
            
            // Update the unit cost display in the header (تكلفة الوحدة)
            const unitPriceHeader = card.querySelector('.item-unit .unit-price-value');
            console.log('Header unit price element:', unitPriceHeader);
            if (unitPriceHeader) {
                unitPriceHeader.textContent = this.formatNumber(adjustedUnitCost);
                console.log('Updated header unit price to:', this.formatNumber(adjustedUnitCost));
            } else {
                console.warn('Header unit price element not found');
            }
            
            // Update the unit cost display in the body (تكلفة الوحدة)
            const unitPriceBody = card.querySelector('.display-group .item-unit .unit-price-value');
            console.log('Body unit price element:', unitPriceBody);
            if (unitPriceBody) {
                unitPriceBody.textContent = this.formatNumber(adjustedUnitCost);
                console.log('Updated body unit price to:', this.formatNumber(adjustedUnitCost));
            } else {
                console.warn('Body unit price element not found');
            }
            
            // Update the selling price display in the header
            const sellPriceHeader = card.querySelector('.sell-price-header');
            if (sellPriceHeader) {
                sellPriceHeader.textContent = this.formatNumber(sellPrice);
            }
            
            // Update the selling price display in the body
            const sellPriceValue = card.querySelector('.sell-price-value');
            if (sellPriceValue) {
                sellPriceValue.textContent = this.formatNumber(sellPrice);
            }
            
            // Update the total display in the header
            const totalElement = card.querySelector('.item-total');
            if (totalElement) {
                totalElement.innerHTML = `الإجمالي: ${this.formatNumber(cardTotal)} جنيه`;
            }
            
            // Update the total display in the body
            const bodyTotalElement = card.querySelector('.body-total-value');
            if (bodyTotalElement) {
                bodyTotalElement.textContent = this.formatNumber(cardTotal);
            }
            
            // Update the card dataset
            card.dataset.sellPrice = sellPrice;
            
            console.log('=== Card display update completed ===');
            
        } catch (error) {
            console.error('Error updating card total display:', error);
        }
    }

    // Add to undo stack
    addToUndoStack(actionData) {
        this.undoStack.push(actionData);
        
        // Keep only last N actions
        if (this.undoStack.length > this.maxUndoActions) {
            this.undoStack.shift();
        }
        
        // Update undo button text
        this.updateUndoButtonText();
    }

    // Get card position for undo
    getCardPosition(card) {
        const cards = Array.from(this.summaryCards.querySelectorAll('.summary-card'));
        return cards.indexOf(card);
    }

    // Undo last action
    undoLastDelete() {
        if (this.undoStack.length === 0) return;
        
        const lastAction = this.undoStack.pop();
        
        if (lastAction.action === 'delete') {
            // Restore the card
            if (lastAction.position >= 0 && lastAction.position < this.summaryCards.children.length) {
                // Insert at specific position
                const targetCard = this.summaryCards.children[lastAction.position];
                if (targetCard) {
                    this.summaryCards.insertBefore(lastAction.cardElement, targetCard);
                } else {
                    this.summaryCards.appendChild(lastAction.cardElement);
                }
            } else {
                // Append to end if position is invalid
                this.summaryCards.appendChild(lastAction.cardElement);
            }
            
            // Save project items FIRST
            this.saveProjectItemsFromDOM();
            
            // Then update all totals and displays
            this.updateSummaryTotal();
            this.updateSummarySellingTotal();
            this.updateSummaryFinalTotal();
            this.updateResourcesSection();
            
            // Update resources totals ribbon after data is saved
            this.updateResourcesTotals();
        }
        
        // Update undo button text
        this.updateUndoButtonText();
        
        // Hide undo button if no more actions
        if (this.undoStack.length === 0 && this.summaryUndoBtn) {
            this.summaryUndoBtn.style.display = 'none';
        }
    }

    // Update undo button text to show count
    updateUndoButtonText() {
        if (this.summaryUndoBtn) {
            const count = this.undoStack.length;
            if (count > 0) {
                this.summaryUndoBtn.textContent = `تراجع (${count})`;
            } else {
                this.summaryUndoBtn.textContent = 'تراجع';
            }
        }
    }

    saveProjectItemsFromDOM() {
        if (!this.currentProjectId || !this.projects[this.currentProjectId]) return;
        const proj = this.projects[this.currentProjectId];
        proj.items = Array.from(this.summaryCards.querySelectorAll('.summary-card')).map(card => card.cardData);
        this.saveProjects();
    }

    updateSummaryTotal() {
        try {
            // Sum all SELECTED cards' unit costs (تكلفة الوحدة) with waste and operation percentages
        let total = 0;
            const selectedCards = this.getSelectedCards();
            
            selectedCards.forEach((card, index) => {
                try {
                    // Use unit price (without tax and risk) for base cost calculation
                    const unitPrice = parseFloat(card.cardData?.unitPrice) || 0;
                    const quantity = parseFloat(card.cardData?.quantity) || 0;
                    const wastePercent = parseFloat(card.cardData?.wastePercent) || 0;
                    const operationPercent = parseFloat(card.cardData?.operationPercent) || 0;
                    
                    // Calculate base cost
                    let baseCost = unitPrice * quantity;
                    
                    // Apply waste and operation percentages (same logic as main calculation)
                    let cardTotal = (baseCost * (1 + wastePercent / 100)) + (baseCost * (operationPercent / 100));
                    
                    total += cardTotal;
                    
                    if (isNaN(cardTotal)) {
                        console.warn(`Card ${index + 1} has invalid calculation:`, { unitPrice, quantity, wastePercent, operationPercent, cardTotal });
                    }
                } catch (cardError) {
                    console.error(`Error processing card ${index + 1}:`, cardError);
                }
            });
            
        if (this.summaryTotal) {
                this.summaryTotal.innerHTML = `<span class="total-value">${this.formatNumber(total)}</span>`;
            } else {
                console.warn('summaryTotal element not found');
            }
            
            // Also update the selling total
            this.updateSummarySellingTotal();
            
        } catch (error) {
            console.error('Error updating summary total:', error);
        }
    }

    updateSummarySellingTotal() {
        try {
            // Sum all SELECTED cards' selling prices (سعر البيع) after risk and tax
            let sellingTotal = 0;
            const selectedCards = this.getSelectedCards();
            
            selectedCards.forEach((card, index) => {
                try {
                    const sell = parseFloat(card.dataset.sellPrice) || parseFloat(card.cardData?.unitPrice) || 0;
                    const quantity = parseFloat(card.cardData?.quantity) || 0;
                    const cardTotal = sell * quantity;
                    sellingTotal += cardTotal;
                    
                    if (isNaN(cardTotal)) {
                        console.warn(`Card ${index + 1} has invalid selling calculation:`, { sell, quantity, cardTotal });
                    }
                } catch (cardError) {
                    console.error(`Error processing selling for card ${index + 1}:`, cardError);
                }
            });
            
            if (this.summarySellingTotal) {
                this.summarySellingTotal.innerHTML = `<span class="selling-total-value">${this.formatNumber(sellingTotal)}</span>`;
            } else {
                console.warn('summarySellingTotal element not found');
            }
            
            // Also update the final total
            this.updateSummaryFinalTotal();
            
        } catch (error) {
            console.error('Error updating selling total:', error);
        }
    }

    updateSummaryFinalTotal() {
        try {
            // Get the selling total from SELECTED cards
            let sellingTotal = 0;
            const selectedCards = this.getSelectedCards();
            
            selectedCards.forEach((card, index) => {
                try {
                    const sellPrice = parseFloat(card.dataset.sellPrice) || parseFloat(card.cardData?.unitPrice) || 0;
                    const quantity = parseFloat(card.cardData?.quantity) || 0;
                    const cardTotal = sellPrice * quantity;
                    sellingTotal += cardTotal;
                    
                    if (isNaN(cardTotal)) {
                        console.warn(`Card ${index + 1} has invalid final calculation:`, { sellPrice, quantity, cardTotal });
                    }
                } catch (cardError) {
                    console.error(`Error processing final for card ${index + 1}:`, cardError);
                }
            });
            
            // Apply supervision percentage
            const supervisionPercent = parseFloat(this.supervisionPercentage?.value) || 0;
            const finalTotal = sellingTotal * (1 + supervisionPercent / 100);
            
            if (this.summaryFinalTotal) {
                this.summaryFinalTotal.innerHTML = `<span class="final-total-value">${this.formatNumber(finalTotal)}</span>`;
            } else {
                console.warn('summaryFinalTotal element not found');
            }
            
        } catch (error) {
            console.error('Error updating final total:', error);
        }
    }

    // Determine the correct unit for a given main/sub item based on user rules
    getCorrectUnit(mainItem, subItem) {
        // Normalize main item: trim and strip leading definite article 'ال'
        const normalize = (s) => (s || '').trim().replace(/^ال\s*/,'');
        // Arabic normalization: remove diacritics/tatweel and unify Alef/Hamza/Yaa/Ta Marbuta
        const normalizeArabic = (s) => (s || '')
            .replace(/[\u064B-\u065F\u0670\u0640]/g, '') // tashkeel + tatweel
            .replace(/[أإآ]/g, 'ا')
            .replace(/ى/g, 'ي')
            .replace(/[ؤئء]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
        const mainNorm = normalize(mainItem);
        const subNorm = (subItem || '').trim();
        const mainNormN = normalizeArabic(mainNorm);
        const subNormN = normalizeArabic(subNorm);
        // Helper for substring match
        const contains = (str, arr) => arr.some(s => str.includes(s));
        const containsN = (str, arr) => arr.some(s => str.includes(normalizeArabic(s)));
        // Special cases for بورسلين
        if (mainItem.includes('بورسلين') || mainNorm.includes('بورسلين') || mainNormN.includes('بورسلين')) {
            if (subNormN.includes('وزر')) return 'مط';
            return 'م2';
        }
        // Special cases for جبسوم بورد
        if (mainItem.includes('جبسوم بورد') || mainNorm.includes('جبسوم بورد') || mainNormN.includes('جبسوم بورد')) {
            const mtSubs = ['ابيض طولي', 'أبيض طولي', 'اخضر طولي', 'أخضر طولي', 'بيوت ستاير', 'بيوت ستائر', 'نور', 'ماجنتك', 'ماجنتك تراك', 'تراك ماجنتك'];
            if (containsN(subNormN, mtSubs)) return 'مط';
            return 'م2';
        }
        // Special cases for تأسيس كهرباء
        if (mainItem.includes('تأسيس كهرباء') || mainNorm.includes('تأسيس كهرباء') || mainNormN.includes('تاسيس كهرباء')) {
            if (subNormN.includes('صواعد')) return 'مط';
            return 'نقطة';
        }
        // removed: تأسيس سباكة
        // Special cases for تأسيس تكييف
        if (mainItem.includes('تأسيس تكييف') || mainNorm.includes('تأسيس تكييف') || mainNorm.includes('تكييفات') || mainNormN.includes('تاسيس تكييف')) return 'مط';
        // Special cases for تأسيس صحي
        if (mainItem.includes('تأسيس صحي') || mainNorm.includes('تأسيس صحي') || mainNormN.includes('تاسيس صحي')) return 'نقطة';
        // Special cases for عزل
        if (mainItem.includes('عزل') || mainNorm.includes('عزل') || mainNormN.includes('عزل')) return 'م2';
        // المتر المسطح الافتراضي لفئات معينة (مع وبدون "ال")
        const m2MainsNormalized = ['مباني','هدم','نقاشة','نقاشه','محارة','رخام'];
        if (m2MainsNormalized.includes(mainNorm) || m2MainsNormalized.includes(mainNormN)) return 'م2';
        // Fallback: try to get from itemsList
        const matchingItems = itemsList.filter(item => item['Main Item'] === mainItem && item['Sub Item'] === subItem);
        let unit = '';
        if (matchingItems.length > 0) {
            const found = matchingItems.find(item => item.Unit && item.Unit !== '');
            unit = found ? found.Unit : (matchingItems[0].Unit || '');
        }
        return unit;
    }

    // Helper to get all resources used in SELECTED items only
    getResourcesSummary() {
        // Map: resourceName -> { type, unit, totalAmount, totalCost, usages: [ {itemTitle, amount, cost, unit} ] }
        const summary = {};
        // For each SELECTED card only
        const selectedCards = this.getSelectedCards();
        selectedCards.forEach(card => {
            // Get item info
            const titleDiv = card.querySelector('.item-title');
            const itemTitle = titleDiv ? titleDiv.textContent : '';
            const quantityInput = card.querySelector('.quantity-input');
            let quantity = 1;
            if (quantityInput) {
                quantity = parseFloat(quantityInput.value) || 0;
            }
            // Read card-level waste/operation percentages to align totals with البنود
            const wasteInput = card.querySelector('.waste-input');
            const operationInput = card.querySelector('.operation-input');
            const wastePercent = wasteInput ? (parseFloat(wasteInput.value) || 0) : 0;
            const operationPercent = operationInput ? (parseFloat(operationInput.value) || 0) : 0;
            const adjustmentFactor = 1 + (wastePercent / 100) + (operationPercent / 100);
            // Find the main/sub item in itemsList
            let mainItem = '', subItem = '';
            if (itemTitle.includes(' - ')) {
                [mainItem, subItem] = itemTitle.split(' - ');
            }
            const matchingItems = itemsList.filter(item => item['Main Item'] === mainItem && item['Sub Item'] === subItem);
            matchingItems.forEach(item => {
                const resource = item.Resource;
                const type = item.Type;
                // Always use the current display unit for UI, but calculations use default unit internally
                const unit = this.getResourceUnit(resource, type) || (item.Unit || '');
                const amount = (parseFloat(item['Quantity per Unit']) || 0) * quantity;
                // Use the active price (respects customPrices and default list prices)
                const pricePerUnit = this.getResourcePrice(resource, type) || 0;
                const baseCost = amount * pricePerUnit;
                const cost = baseCost * adjustmentFactor;
                if (!summary[resource]) {
                    summary[resource] = {
                        type,
                        unit,
                        totalAmount: 0,
                        totalCost: 0,
                        usages: []
                    };
                }
                summary[resource].totalAmount += amount;
                summary[resource].totalCost += cost;
                summary[resource].usages.push({ itemTitle, amount, cost, unit });
            });
        });
        return summary;
    }

    // Replace renderResourcesSummary and updateResourcesSection with new logic for three collapsible panels
    renderResourcesAccordion() {
        // Get summary by resource
        const summary = this.getResourcesSummary();
        // Prepare categorized lists
        const categories = {
            'خامات': [],
            'مصنعيات': [],
            'عمالة': []
        };
        Object.entries(summary).forEach(([resource, data]) => {
            if (categories[data.type]) {
                categories[data.type].push({ resource, data });
            }
        });
        
        // Sort each category by total cost (highest to lowest)
        Object.keys(categories).forEach(category => {
            categories[category].sort((a, b) => {
                const totalCostA = a.data.usages.reduce((sum, u) => sum + u.cost, 0);
                const totalCostB = b.data.usages.reduce((sum, u) => sum + u.cost, 0);
                return totalCostB - totalCostA; // High to low
            });
        });
        
        // Helper to render resource rows for a category
        const renderRows = (list) =>
            list.map(({ resource, data }, index) => {
                // Calculate totals for this resource
                const totalCost = data.usages.reduce((sum, u) => sum + u.cost, 0);
                const totalQuantity = data.usages.reduce((sum, u) => sum + u.amount, 0);
                
                return `<div class="resource-row" data-rank="${index + 1}">
                    <div class="resource-header">
                    <span class="resource-name">${resource}</span>
                        <div class="resource-totals">
                            <span class="total-cost">
                                <span class="label">إجمالي التكلفة:</span>
                                <span class="value">${this.formatNumber(totalCost)} جنيه</span>
                            </span>
                            <span class="total-quantity">
                                <span class="label">إجمالي الكمية:</span>
                                <span class="value">${this.formatNumber(totalQuantity)} ${data.unit}</span>
                            </span>
                        </div>
                    </div>
                    <div class="resource-usage-details">
                        <table class="resource-usage-table">
                            <thead>
                                <tr>
                                    <th>البند</th>
                                    <th>الكمية</th>
                                    <th>التكلفة</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${data.usages.map(u => `
                                    <tr>
                                        <td class="usage-item-title">${u.itemTitle}</td>
                                        <td class="usage-amount">${this.formatNumber(u.amount)} ${u.unit}</td>
                                        <td class="usage-cost">${this.formatNumber(u.cost)} جنيه</td>
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>`;
            }).join('');
        
        // Render each panel with sorting info
        if (this.resourcesMaterialsBody) {
            this.resourcesMaterialsBody.innerHTML = renderRows(categories['خامات']);
        }
        if (this.resourcesWorkmanshipBody) {
            this.resourcesWorkmanshipBody.innerHTML = renderRows(categories['مصنعيات']);
        }
        if (this.resourcesLaborBody) {
            this.resourcesLaborBody.innerHTML = renderRows(categories['عمالة']);
        }
    }

    updateResourcesSection() {
        this.renderResourcesAccordion();
        [
            { header: this.resourcesMaterialsHeader, body: this.resourcesMaterialsBody },
            { header: this.resourcesWorkmanshipHeader, body: this.resourcesWorkmanshipBody },
            { header: this.resourcesLaborHeader, body: this.resourcesLaborBody }
        ].forEach(({ header, body }) => {
            if (header && body) {
                body.style.display = 'none';
                const btn = header.querySelector('.accordion-toggle');
                if (btn) {
                    btn.setAttribute('aria-expanded', 'false');
                    btn.textContent = '+';
                    btn.onclick = (e) => {
                        e.stopPropagation();
                        const expanded = body.style.display === 'block';
                        body.style.display = expanded ? 'none' : 'block';
                        btn.setAttribute('aria-expanded', expanded ? 'false' : 'true');
                        btn.textContent = expanded ? '+' : '–';
                    };
                }
                header.onclick = (e) => {
                    if (e.target.classList.contains('accordion-toggle')) return;
                    const btn = header.querySelector('.accordion-toggle');
                    btn && btn.click();
                };
            }
        });
        
        // Update totals after rendering the accordion
        this.updateResourcesTotals();
    }

    // --- Project Management Logic ---
    setupProjectManagement() {
        this.projectForm.onsubmit = (e) => {
            e.preventDefault();
            const name = this.projectNameInput.value.trim();
            const code = this.projectCodeInput.value.trim();
            const type = this.projectTypeInput.value;
            const area = parseFloat(this.projectAreaInput.value);
            const floor = parseInt(this.projectFloorInput.value);
            if (!name || !code || !type || !area || !floor) return;
            // Create new project object
            const id = 'proj_' + Date.now();
            this.projects[id] = {
                name, code, type, area, floor,
                prices: { materials: {}, workmanship: {}, labor: {} },
                items: [],
                laborExtrasPerFloor: {},
                laborFloorLevel: 1
            };
            this.saveProjects();
            this.currentProjectId = id;
            this.saveCurrentProjectId();
            this.renderProjectsList();
            this.loadProjectData();
            this.projectForm.reset();
        };
        
        // Remove syncing labor floor from project floor; only save project floor and refresh display
        if (this.projectFloorInput) {
            const syncFloor = () => {
                const floorVal = parseInt(this.projectFloorInput.value) || 1;
                if (this.currentProjectId && this.projects[this.currentProjectId]) {
                    this.projects[this.currentProjectId].floor = floorVal;
                    this.saveProjects();
                    this.renderCurrentProjectDisplay();
                }
            };
            this.projectFloorInput.addEventListener('input', syncFloor);
            this.projectFloorInput.addEventListener('change', syncFloor);
        }
        
        this.renderProjectsList();
    }

    renderProjectsList() {
        this.projectsList.innerHTML = '';
        Object.entries(this.projects).forEach(([id, proj]) => {
            const div = document.createElement('div');
            div.className = 'project-list-item';
            div.innerHTML = `
                <span class="project-name">${proj.name}</span>
                <span class="project-code">(${proj.code})</span>
                <button data-id="${id}" class="select-btn project-action-btn">تحديد</button>
                <button data-id="${id}" class="delete-btn project-action-btn">حذف</button>
            `;
            // Select
            div.querySelector('.select-btn').onclick = () => {
                this.currentProjectId = id;
                this.saveCurrentProjectId();
                this.loadProjectData();
                this.renderProjectsList();
            };
            // Delete
            div.querySelector('.delete-btn').onclick = () => {
                if (confirm('هل أنت متأكد من حذف هذا المشروع؟')) {
                    delete this.projects[id];
                    if (this.currentProjectId === id) {
                        this.currentProjectId = Object.keys(this.projects)[0] || null;
                        this.saveCurrentProjectId();
                    }
                    this.saveProjects();
                    this.renderProjectsList();
                    this.loadProjectData();
                }
            };
            if (this.currentProjectId === id) {
                div.style.background = 'var(--bg-tertiary)';
            }
            this.projectsList.appendChild(div);
        });
        // Show current project
        this.renderCurrentProjectDisplay();
    }

    renderCurrentProjectDisplay() {
        const proj = this.projects[this.currentProjectId];
        if (proj) {
            this.currentProjectDisplay.innerHTML = `
                <div class="current-project-main">
                    <span class="project-name">${proj.name}</span>
                    <span class="project-code">(${proj.code})</span>
                </div>
                <div class="current-project-sub">
                    <span class="project-type"><span class="label">النوع:</span> <span class="value">${proj.type}</span></span>
                    <span class="sep">•</span>
                    <span class="project-area"><span class="label">المساحة:</span> <span class="value">${this.formatNumber(proj.area)} م²</span></span>
                    <span class="sep">•</span>
                    <span class="project-floor"><span class="label">الأدوار:</span> <span class="value">${proj.floor}</span></span>
                </div>
            `;
        } else {
            this.currentProjectDisplay.innerHTML = '<span>لا يوجد مشروع محدد</span>';
        }
    }

    // --- Project Data Storage ---
    loadProjects() {
        try {
            return JSON.parse(localStorage.getItem('projects')) || {};
        } catch {
            return {};
        }
    }
    saveProjects() {
        localStorage.setItem('projects', JSON.stringify(this.projects));
    }
    loadCurrentProjectId() {
        return localStorage.getItem('currentProjectId') || Object.keys(this.projects)[0] || null;
    }
    saveCurrentProjectId() {
        localStorage.setItem('currentProjectId', this.currentProjectId || '');
    }

    // --- Project Data Context ---
    loadProjectData() {
        // Clear undo stack when loading new project
        this.clearUndoStack();
        
        // Show current project info
        this.renderCurrentProjectDisplay();
        // If no project, clear UI
        if (!this.currentProjectId || !this.projects[this.currentProjectId]) {
            this.clearProjectUI();
            return;
        }
        // Load prices and items for this project
        const proj = this.projects[this.currentProjectId];
        // 1. Prices
        this.customPrices = new Map(Object.entries(proj.prices && proj.prices.customPrices || {}));
        this.customUnits = new Map(Object.entries(proj.prices && proj.prices.customUnits || {}));
        // 1.b Labor extras per floor and labor floor level
        this.laborExtrasPerFloor = Object.assign({}, proj.laborExtrasPerFloor || {});
        this.laborFloorLevel = proj.laborFloorLevel || 1;
        // 2. Summary items
        this.summaryCards.innerHTML = '';
        (proj.items || [])
            .filter(cardData => cardData.mainItem !== 'تأسيس سباكة')
            .forEach(cardData => this.addSummaryCard(cardData));
        
        // Ensure all loaded cards have proper sellPrice in dataset
        this.ensureAllCardsHaveSellPrice();
        
        // After loading cards, migrate old cards to current pricing logic
        this.recalculateAllCardsAndSave();
        
        // Update selection count after loading all cards
        this.updateSelectionCount();
        
        // Ensure selection system is properly set up
        this.setupSelectionSystem();
        
        // Update all loaded cards with the restored global percentages
        this.updateAllSummaryCardsWithNewPercentages();
        
        // 3. Labor floor level
        if (this.laborFloorLevelInput) {
            this.laborFloorLevelInput.value = proj.floor || 1;
            this.updateLaborPricesForFloor();
        }
        // 4. Reset calculator inputs
        if (this.mainItemSelect) this.mainItemSelect.value = '';
        if (this.subItemSelect) {
            this.subItemSelect.value = '';
            this.subItemSelect.disabled = true;
        }
        if (this.quantityInput) this.quantityInput.value = '';
        
        // Restore saved waste and operation percentages or use defaults
        if (this.wastePercentInput) {
            this.wastePercentInput.value = proj.globalWastePercent || 0;
        }
        if (this.operationPercentInput) {
            this.operationPercentInput.value = proj.globalOperationPercent || 0;
        }
        
        if (this.unitPriceDisplay) this.unitPriceDisplay.textContent = '';
        if (this.totalCostElement) this.totalCostElement.textContent = '0.00';
        if (this.resultsSection) this.resultsSection.style.display = 'none';
        // 5. Prices section
        this.loadPricesSection();
        // 6. Recalculate and update all UI
        this.calculate();
        this.updateSummaryTotal();
        this.updateResourcesSection();
        // Load custom rates for this project
        this.customRates = (proj.customRates || {});
    }

    // Clear undo stack
    clearUndoStack() {
        this.undoStack = [];
        this.updateUndoButtonText();
        if (this.summaryUndoBtn) {
            this.summaryUndoBtn.style.display = 'none';
        }
    }

    clearProjectUI() {
        this.summaryCards.innerHTML = '';
        if (this.laborFloorLevelInput) this.laborFloorLevelInput.value = 1;
        if (this.mainItemSelect) this.mainItemSelect.value = '';
        if (this.subItemSelect) {
            this.subItemSelect.value = '';
            this.subItemSelect.disabled = true;
        }
        if (this.quantityInput) this.quantityInput.value = '';
        if (this.wastePercentInput) this.wastePercentInput.value = '';
        if (this.operationPercentInput) this.operationPercentInput.value = '';
        if (this.unitPriceDisplay) this.unitPriceDisplay.textContent = '';
        if (this.totalCostElement) this.totalCostElement.textContent = '0.00';
        if (this.resultsSection) this.resultsSection.style.display = 'none';
        this.loadPricesSection();
        this.updateSummaryTotal();
        this.updateResourcesSection();
        this.calculate();
        
        // Clear undo stack
        this.clearUndoStack();
    }

    // Override price and unit setters to save to project
    setCustomPrice(resourceName, newPrice) {
        this.customPrices.set(resourceName, newPrice);
        this.saveProjectPrices();
    }
    setCustomUnit(resourceName, newUnit) {
        this.customUnits.set(resourceName, newUnit);
        this.saveProjectPrices();
    }
    saveProjectPrices() {
        if (!this.currentProjectId || !this.projects[this.currentProjectId]) return;
        this.projects[this.currentProjectId].prices = {
            customPrices: Object.fromEntries(this.customPrices),
            customUnits: Object.fromEntries(this.customUnits)
        };
        this.saveProjects();
    }

    saveProjectCustomRates() {
        if (!this.currentProjectId || !this.projects[this.currentProjectId]) return;
        this.projects[this.currentProjectId].customRates = this.customRates;
        this.saveProjects();
    }

    saveProjectLaborExtras() {
        if (!this.currentProjectId || !this.projects[this.currentProjectId]) return;
        this.projects[this.currentProjectId].laborExtrasPerFloor = this.laborExtrasPerFloor;
        this.saveProjects();
    }
    saveProjectLaborFloorLevel() {
        if (!this.currentProjectId || !this.projects[this.currentProjectId]) return;
        
        this.projects[this.currentProjectId].laborFloorLevel = this.laborFloorLevel;
        this.saveProjects();
        
        console.log('Saved labor floor level to project:', this.laborFloorLevel);
    }

    // Helper to format rates to up to 4 decimal places, removing trailing zeros
    formatRate(val) {
        return parseFloat(val).toFixed(4).replace(/\.0+$|0+$/,'').replace(/\.$/, '');
    }

    // Show item details modal for a summary card
    showItemDetailsModal(cardData) {
        if (!this.itemDetailsModal) return;
        // Set modal title
        this.itemDetailsTitle.textContent = `تفاصيل البند: ${cardData.mainItem} - ${cardData.subItem}`;
        // Render cost details for this item
        this.itemDetailsContent.innerHTML = this.renderItemCostDetails(cardData);
        this.itemDetailsModal.classList.add('show');
        this.itemDetailsModal.style.display = 'flex';
    }
    hideItemDetailsModal() {
        if (this.itemDetailsModal) {
            this.itemDetailsModal.classList.remove('show');
            this.itemDetailsModal.style.display = 'none';
        }
    }

    // Render the cost details breakdown for a summary card (returns HTML)
    renderItemCostDetails(cardData) {
        // Find the matching items for this card
        const matchingItems = itemsList.filter(item =>
            item['Main Item'] === cardData.mainItem &&
            item['Sub Item'] === cardData.subItem
        );
        if (matchingItems.length === 0) return '<div>لا توجد تفاصيل لهذا البند.</div>';
        // Use the card's quantity and current project custom rates
        const quantity = cardData.quantity;
        const customRates = this.customRates || {};
        // Group by type
        const groups = { 'خامات': [], 'مصنعيات': [], 'عمالة': [] };
        matchingItems.forEach(item => {
            const resourceInfo = this.getResourceInfo(item.Resource, item.Type);
            if (resourceInfo) {
                const unitPrice = this.getResourcePrice(item.Resource, item.Type);
                const rateKey = `${cardData.mainItem}||${cardData.subItem}||${item.Resource}`;
                const itemQuantity = customRates[rateKey] !== undefined ? parseFloat(customRates[rateKey]) : (parseFloat(item['Quantity per Unit']) || 0);
                const totalQuantity = itemQuantity * quantity;
                const safeUnitPrice = isNaN(unitPrice) ? 0 : unitPrice;
                const totalCost = safeUnitPrice * totalQuantity;
                const displayUnit = this.getResourceUnit(item.Resource, item.Type);
                groups[item.Type].push({
                    resource: item.Resource,
                    rate: this.formatRate(itemQuantity),
                    quantity: totalQuantity,
                    unit: displayUnit,
                    unitPrice: safeUnitPrice,
                    totalCost: totalCost
                });
            }
        });
        // Helper to render a group table
        const renderTable = (arr) => {
            if (!arr.length) return '';
            let total = arr.reduce((sum, x) => sum + x.totalCost, 0);
            return `
                <table>
                    <thead>
                        <tr>
                            <th>المورد</th>
                            <th>معدل الاستخدام</th>
                            <th>الكمية المطلوبة</th>
                            <th>الوحدة</th>
                            <th>سعر الوحدة</th>
                            <th>التكلفة</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${arr.map(x => `
                            <tr>
                                <td>${x.resource}</td>
                                <td>${x.rate}</td>
                                <td>${this.formatNumber(x.quantity)} ${x.unit}</td>
                                <td>${x.unit}</td>
                                <td>${this.formatNumber(x.unitPrice)} جنيه</td>
                                <td>${this.formatNumber(x.totalCost)} جنيه</td>
                            </tr>
                        `).join('')}
                    </tbody>
                    <tfoot>
                        <tr style="font-weight:bold;">
                            <td colspan="5">الإجمالي</td>
                            <td>${this.formatNumber(total)} جنيه</td>
                        </tr>
                    </tfoot>
                </table>
            `;
        };
        // Accordion items
        const accordionItems = [
            { key: 'خامات', label: 'الخامات' },
            { key: 'مصنعيات', label: 'المصنعيات' },
            { key: 'عمالة', label: 'العمالة' }
        ].map(({key, label}, idx) => {
            const hasData = groups[key].length > 0;
            return `
                <div class="accordion-item">
                    <div class="accordion-header" data-idx="${idx}">
                        <span class="section-title">${label}</span>
                        <button class="accordion-toggle" aria-expanded="false">+</button>
                    </div>
                    <div class="accordion-body" style="display:none;">
                        ${hasData ? renderTable(groups[key]) : '<div style="color:#888;">لا توجد بيانات</div>'}
                    </div>
                </div>
            `;
        }).join('');
        // Modal accordion wrapper
        setTimeout(() => {
            // Add expand/collapse logic after modal content is rendered
            const modal = document.getElementById('itemDetailsModal');
            if (!modal) return;
            modal.querySelectorAll('.modal-accordion .accordion-header').forEach(header => {
                const btn = header.querySelector('.accordion-toggle');
                const body = header.parentElement.querySelector('.accordion-body');
                if (btn && body) {
                    btn.onclick = (e) => {
                        e.stopPropagation();
                        const expanded = body.style.display === 'block';
                        body.style.display = expanded ? 'none' : 'block';
                        btn.setAttribute('aria-expanded', expanded ? 'false' : 'true');
                        btn.textContent = expanded ? '+' : '–';
                    };
                    header.onclick = (e) => {
                        if (e.target.classList.contains('accordion-toggle')) return;
                        btn.click();
                    };
                    // Collapsed by default
                    body.style.display = 'none';
                    btn.setAttribute('aria-expanded', 'false');
                    btn.textContent = '+';
                }
            });
        }, 0);
        return `<div class="modal-accordion">${accordionItems}</div>`;
    }




    exportProjectToExcel() {
        try {
            // Check if XLSX library is loaded
            if (typeof XLSX === 'undefined') {
                alert('مكتبة Excel غير محملة. يرجى التأكد من تحميل الصفحة بشكل صحيح.');
            return;
        }

            // Test basic XLSX functionality
            try {
                const testWb = XLSX.utils.book_new();
                const testData = [['Test', 'Data']];
                const testSheet = XLSX.utils.aoa_to_sheet(testData);
                XLSX.utils.book_append_sheet(testWb, testSheet, 'Test');
                console.log('Basic XLSX functionality test passed');
            } catch (testError) {
                console.error('Basic XLSX functionality test failed:', testError);
                alert('مشكلة في مكتبة Excel. يرجى تحديث الصفحة والمحاولة مرة أخرى.');
                return;
            }

        // Get current project
        const proj = this.projects[this.currentProjectId];
        if (!proj) {
            alert('لا يوجد مشروع محدد.');
            return;
        }

            // Check if required elements exist
            if (!this.summaryCards) {
                alert('عناصر الصفحة غير جاهزة. يرجى المحاولة مرة أخرى.');
                return;
            }

            console.log('Starting Excel export for project:', proj.name);

            // Create workbook
            const wb = XLSX.utils.book_new();
            
            // Set workbook-level RTL
            wb.Workbook = {
                Views: [
                    {
                        RTL: true
                    }
                ]
            };

            // Get the resources summary data once for all sheets
        const resourcesSummary = this.getResourcesSummary();
            console.log('Resources summary for export:', resourcesSummary);
            
            if (!resourcesSummary || Object.keys(resourcesSummary).length === 0) {
                console.warn('No resources data found for export');
            }

            // 1. Project Overview Sheet (معلومات المشروع)
            const projectData = [
                ['معلومات المشروع', ''],
                ['اسم المشروع', proj.name],
                ['كود المشروع', proj.code],
                ['نوع المشروع', proj.type],
                ['المساحة', `${this.formatNumber(proj.area)} م²`],
                ['عدد الأدوار', proj.floor],
                ['', ''],
                ['تاريخ التصدير', new Date().toISOString().split('T')[0]],
                ['وقت التصدير', new Date().toTimeString().split(' ')[0]]
            ];

            const projectSheet = XLSX.utils.aoa_to_sheet(projectData);
            projectSheet['!rtl'] = true;
            
            // Set column widths for project sheet
            projectSheet['!cols'] = [
                { width: 25 },
                { width: 35 }
            ];

            XLSX.utils.book_append_sheet(wb, projectSheet, 'معلومات المشروع');
            console.log('Project overview sheet added');

            // 2. Materials Sheet (الخامات)
            if (this.resourcesMaterialsBody) {
                console.log('Processing materials sheet...');
                
                const materialsData = [
                    ['إدارة الموارد - الخامات', '', '', '', '', ''],
                    ['', '', '', '', '', ''],
                    ['', '', '', '', '', ''],
                    ['اسم المورد', 'الوحدة', 'الكمية', 'سعر الوحدة (جنيه)', 'التكلفة الإجمالية (جنيه)', 'ملاحظات'],
                    ['', '', '', '', '']
                ];

                // Get the actual data from the summary instead of parsing DOM
        const resourcesSummary = this.getResourcesSummary();
                const materialsResources = Object.entries(resourcesSummary).filter(([resource, data]) => data.type === 'خامات');
                
                console.log('Materials resources from summary:', materialsResources);
                
                materialsResources.forEach(([resource, data]) => {
                    try {
                        const resourceName = resource;
                        const unit = data.unit || '';
                        const actualQuantity = this.formatNumber(data.totalAmount);
                        
                        // Calculate unit price from total cost / total amount
                        let unitPrice = '';
                        if (data.totalAmount > 0) {
                            unitPrice = this.formatNumber(data.totalCost / data.totalAmount);
                        }
                        
                        const totalCost = this.formatNumber(data.totalCost) + ' جنيه';

                        console.log(`Adding material: ${resourceName}, unit: ${unit}, quantity: ${actualQuantity}, unitPrice: ${unitPrice}, totalCost: ${totalCost}`);
                        
                        materialsData.push([resourceName, unit, actualQuantity, unitPrice, totalCost, '']);
                    } catch (error) {
                        console.error('Error processing material resource:', error);
                    }
                });

                // Add totals section
                const materialsTotal = this.calculateSectionTotal(this.resourcesMaterialsBody);
                materialsData.push(['', '', '', '', '', '']);
                materialsData.push(['', '', '', '', '', '']);
                materialsData.push(['إجمالي تكلفة الخامات', '', '', '', materialsTotal, '']);
                materialsData.push(['', '', '', '', '', '']);
                materialsData.push(['ملاحظات', 'تشمل جميع المواد الأساسية المطلوبة للمشروع', '', '', '', '']);

                const materialsSheet = XLSX.utils.aoa_to_sheet(materialsData);
                materialsSheet['!rtl'] = true;
                materialsSheet['!cols'] = [
                    { width: 35 }, // اسم المورد
                    { width: 15 }, // الوحدة
                    { width: 15 }, // الكمية
                    { width: 25 }, // سعر الوحدة
                    { width: 30 }, // التكلفة الإجمالية
                    { width: 25 }  // ملاحظات
                ];

                XLSX.utils.book_append_sheet(wb, materialsSheet, 'الخامات');
                console.log('Materials sheet added');
            }

            // 3. Workmanship Sheet (المصنعيات)
            if (this.resourcesWorkmanshipBody) {
                console.log('Processing workmanship sheet...');
                
                const workmanshipData = [
                    ['إدارة الموارد - المصنعيات', '', '', '', '', ''],
                    ['', '', '', '', '', ''],
                    ['', '', '', '', '', ''],
                    ['اسم المورد', 'الوحدة', 'الكمية', 'سعر الوحدة (جنيه)', 'التكلفة الإجمالية (جنيه)', 'ملاحظات'],
                    ['', '', '', '', '']
                ];

                // Get the actual data from the summary instead of parsing DOM
                const workmanshipResources = Object.entries(resourcesSummary).filter(([resource, data]) => data.type === 'مصنعيات');
                
                console.log('Workmanship resources from summary:', workmanshipResources);
                
                workmanshipResources.forEach(([resource, data]) => {
                    try {
                        const resourceName = resource;
                        const unit = data.unit || '';
                        const actualQuantity = this.formatNumber(data.totalAmount);
                        
                        // Calculate unit price from total cost / total amount
                        let unitPrice = '';
                        if (data.totalAmount > 0) {
                            unitPrice = this.formatNumber(data.totalCost / data.totalAmount);
                        }
                        
                        const totalCost = this.formatNumber(data.totalCost) + ' جنيه';

                        console.log(`Adding workmanship: ${resourceName}, unit: ${unit}, quantity: ${actualQuantity}, unitPrice: ${unitPrice}, totalCost: ${totalCost}`);
                        
                        workmanshipData.push([resourceName, unit, actualQuantity, unitPrice, totalCost, '']);
                    } catch (error) {
                        console.error('Error processing workmanship resource:', error);
                    }
                });

                // Add totals section
                const workmanshipTotal = this.calculateSectionTotal(this.resourcesWorkmanshipBody);
                workmanshipData.push(['', '', '', '', '', '']);
                workmanshipData.push(['', '', '', '', '', '']);
                workmanshipData.push(['إجمالي تكلفة المصنعيات', '', '', '', workmanshipTotal, '']);
                workmanshipData.push(['', '', '', '', '', '']);
                workmanshipData.push(['ملاحظات', 'تشمل جميع المصنعيات والمنتجات الجاهزة', '', '', '', '']);

                const workmanshipSheet = XLSX.utils.aoa_to_sheet(workmanshipData);
                workmanshipSheet['!rtl'] = true;
                workmanshipSheet['!cols'] = [
                    { width: 35 }, // اسم المورد
                    { width: 15 }, // الوحدة
                    { width: 15 }, // الكمية
                    { width: 25 }, // سعر الوحدة
                    { width: 30 }, // التكلفة الإجمالية
                    { width: 25 }  // ملاحظات
                ];

                XLSX.utils.book_append_sheet(wb, workmanshipSheet, 'المصنعيات');
                console.log('Workmanship sheet added');
            }

            // 4. Labor Sheet (العمالة)
            if (this.resourcesLaborBody) {
                console.log('Processing labor sheet...');
                
                const laborData = [
                    ['إدارة الموارد - العمالة', '', '', '', '', ''],
                    ['', '', '', '', '', ''],
                    ['', '', '', '', '', ''],
                    ['اسم المورد', 'الوحدة', 'الكمية', 'سعر الوحدة (جنيه)', 'التكلفة الإجمالية (جنيه)', 'ملاحظات'],
                    ['', '', '', '', '']
                ];

                // Get the actual data from the summary instead of parsing DOM
                const laborResources = Object.entries(resourcesSummary).filter(([resource, data]) => data.type === 'عمالة');
                
                console.log('Labor resources from summary:', laborResources);
                
                laborResources.forEach(([resource, data]) => {
                    try {
                        const resourceName = resource;
                        const unit = data.unit || '';
                        const actualQuantity = this.formatNumber(data.totalAmount);
                        
                        // Calculate unit price from total cost / total amount
                        let unitPrice = '';
                        if (data.totalAmount > 0) {
                            unitPrice = this.formatNumber(data.totalCost / data.totalAmount);
                        }
                        
                        const totalCost = this.formatNumber(data.totalCost) + ' جنيه';

                        console.log(`Adding labor: ${resourceName}, unit: ${unit}, quantity: ${actualQuantity}, unitPrice: ${unitPrice}, totalCost: ${totalCost}`);
                        
                        laborData.push([resourceName, unit, actualQuantity, unitPrice, totalCost, '']);
                    } catch (error) {
                        console.error('Error processing labor resource:', error);
                    }
                });

                // Add totals section
                const laborTotal = this.calculateSectionTotal(this.resourcesLaborBody);
                laborData.push(['', '', '', '', '', '']);
                laborData.push(['', '', '', '', '', '']);
                laborData.push(['إجمالي تكلفة العمالة', '', '', '', laborTotal, '']);
                laborData.push(['', '', '', '', '', '']);
                laborData.push(['ملاحظات', 'تشمل جميع خدمات العمالة والتنفيذ', '', '', '', '']);

                const laborSheet = XLSX.utils.aoa_to_sheet(laborData);
                laborSheet['!rtl'] = true;
                laborSheet['!cols'] = [
                    { width: 35 }, // اسم المورد
                    { width: 15 }, // الوحدة
                    { width: 15 }, // الكمية
                    { width: 25 }, // سعر الوحدة
                    { width: 30 }, // التكلفة الإجمالية
                    { width: 25 }  // ملاحظات
                ];

                XLSX.utils.book_append_sheet(wb, laborSheet, 'العمالة');
                console.log('Labor sheet added');
            }

            // 5. Summary Sheet (الملخص)
            console.log('Processing summary sheet...');
            const summaryData = [
                ['ملخص المشروع - البنود', '', '', '', '', ''],
                ['', '', '', '', '', ''],
                ['', '', '', '', '', ''],
                ['البند الرئيسي', 'البند الفرعي', 'الكمية', 'الوحدة', 'سعر الوحدة (جنيه)', 'التكلفة الإجمالية (جنيه)'],
                ['', '', '', '', '', '']
            ];

            // Add SELECTED summary cards data only
            const selectedCards = this.getSelectedCards();
            console.log('Found selected summary cards:', selectedCards.length);
            
            selectedCards.forEach(card => {
                try {
                    const cardData = card.cardData;
                    if (cardData) {
                        summaryData.push([
                            cardData.mainItem || '',
                            cardData.subItem || '',
                            this.formatNumber(cardData.quantity) || '',
                            cardData.unit || '',
                            this.formatNumber(cardData.unitPrice) || '',
                            this.formatNumber(cardData.total) || ''
                        ]);
                    }
                } catch (error) {
                    console.error('Error processing summary card:', error);
                }
            });

            // Add summary totals
            const summaryTotal = this.calculateSummaryTotal();
            const summarySellingTotal = this.calculateSummarySellingTotal();
            const summaryFinalTotal = this.calculateSummaryFinalTotal();

            summaryData.push(['', '', '', '', '', '']);
            summaryData.push(['', '', '', '', '', '']);
            summaryData.push(['إجمالي التكلفة الأساسية', '', '', '', '', summaryTotal]);
            summaryData.push(['إجمالي سعر البيع', '', '', '', '', summarySellingTotal]);
            summaryData.push(['إجمالي سعر البيع النهائي', '', '', '', '', summaryFinalTotal]);
            summaryData.push(['', '', '', '', '', '']);
            summaryData.push(['ملاحظات', 'تشمل جميع بنود المشروع مع حسابات المخاطر والضرائب', '', '', '', '']);

            const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
            summarySheet['!rtl'] = true;
            summarySheet['!cols'] = [
                { width: 25 }, // البند الرئيسي
                { width: 25 }, // البند الفرعي
                { width: 15 }, // الكمية
                { width: 15 }, // الوحدة
                { width: 25 }, // سعر الوحدة
                { width: 30 }  // التكلفة الإجمالية
            ];

            XLSX.utils.book_append_sheet(wb, summarySheet, 'الملخص');
            console.log('Summary sheet added');

            // 6. Totals Overview Sheet (الإجماليات)
            console.log('Processing totals sheet...');
            const totalsData = [
                ['إجماليات المشروع - ملخص شامل', ''],
                ['', ''],
                ['', ''],
                ['تفاصيل التكاليف', 'المبلغ (جنيه)'],
                ['', ''],
                ['إجمالي تكلفة الخامات', this.getSectionTotalDisplay('resourcesMaterialsTotal')],
                ['إجمالي تكلفة المصنعيات', this.getSectionTotalDisplay('resourcesWorkmanshipTotal')],
                ['إجمالي تكلفة العمالة', this.getSectionTotalDisplay('resourcesLaborTotal')],
                ['', ''],
                ['المجموع الكلي للموارد', this.getSectionTotalDisplay('resourcesGrandTotal')],
                ['', ''],
                ['إجمالي التكلفة الأساسية', summaryTotal],
                ['إجمالي سعر البيع', summarySellingTotal],
                ['إجمالي سعر البيع النهائي', summaryFinalTotal],
                ['', ''],
                ['ملاحظات', 'تم حساب جميع التكاليف بناءً على البيانات المدخلة في النظام']
            ];

            const totalsSheet = XLSX.utils.aoa_to_sheet(totalsData);
            totalsSheet['!rtl'] = true;
            totalsSheet['!cols'] = [
                { width: 40 },
                { width: 30 }
            ];

            XLSX.utils.book_append_sheet(wb, totalsSheet, 'الإجماليات');
            console.log('Totals sheet added');

            // Set default row height for all sheets
            if (wb.Sheets && typeof wb.Sheets === 'object') {
                Object.keys(wb.Sheets).forEach(sheetName => {
                    if (wb.Sheets[sheetName]) {
                        wb.Sheets[sheetName]['!rows'] = Array(50).fill({ hpt: 20 });
                    }
                });
            }

            console.log('All sheets prepared, starting download...');

            // Download the file
            const fileName = `${proj.name || 'مشروع'}_${new Date().toISOString().split('T')[0]}.xlsx`;
            XLSX.writeFile(wb, fileName);
            
            console.log('File downloaded successfully:', fileName);
            
            // Success message
            alert(`تم تصدير المشروع بنجاح إلى ملف: ${fileName}`);

        } catch (error) {
            console.error('Error exporting to Excel:', error);
            console.error('Error details:', {
                message: error.message,
                stack: error.stack,
                name: error.name
            });
            
            // Check specific conditions that might cause failure
            console.log('Debug info:', {
                hasXLSX: typeof XLSX !== 'undefined',
                hasProject: !!this.projects[this.currentProjectId],
                hasSummaryCards: !!this.summaryCards,
                hasMaterialsBody: !!this.resourcesMaterialsBody,
                hasWorkmanshipBody: !!this.resourcesWorkmanshipBody,
                hasLaborBody: !!this.resourcesLaborBody
            });
            
            alert('حدث خطأ أثناء التصدير إلى Excel. يرجى المحاولة مرة أخرى.\n\nالتفاصيل: ' + error.message);
        }
    }

    // Export Resources Management section to HTML with three separate files
    exportResourcesToHtml() {
        try {
        // Get current project
        const proj = this.projects[this.currentProjectId];
        if (!proj) {
            alert('لا يوجد مشروع محدد.');
            return;
        }

            console.log('Starting Resources HTML export for project:', proj.name);

            // Get the resources summary data
        const resourcesSummary = this.getResourcesSummary();
            console.log('Resources summary for export:', resourcesSummary);
            
            if (!resourcesSummary || Object.keys(resourcesSummary).length === 0) {
                alert('لا توجد بيانات موارد للتصدير.');
                return;
            }

            // Export Materials (الخامات)
            this.exportResourceTypeToHtml('خامات', resourcesSummary, proj, 'الخامات');
            
            // Export Workmanship (المصنعيات)
            this.exportResourceTypeToHtml('مصنعيات', resourcesSummary, proj, 'المصنعيات');
            
            // Export Labor (العمالة)
            this.exportResourceTypeToHtml('عمالة', resourcesSummary, proj, 'العمالة');

            // Success message
            alert('تم تصدير إدارة الموارد بنجاح إلى 3 ملفات HTML منفصلة!');

        } catch (error) {
            console.error('Error exporting resources to HTML:', error);
            alert('حدث خطأ أثناء التصدير إلى HTML. يرجى المحاولة مرة أخرى.\n\nالتفاصيل: ' + error.message);
        }
    }

    // Export specific resource type to HTML
    exportResourceTypeToHtml(resourceType, resourcesSummary, proj, sheetName) {
        try {
            // Get resources of this type with full details
            const typeResources = Object.entries(resourcesSummary).filter(([resource, data]) => data.type === resourceType);
            
            // Sort by total cost (highest to lowest) - same as display
            typeResources.sort((a, b) => {
                const totalCostA = a[1].usages.reduce((sum, u) => sum + u.cost, 0);
                const totalCostB = b[1].usages.reduce((sum, u) => sum + u.cost, 0);
                return totalCostB - totalCostA;
            });

            // Calculate total for this resource type
            const resourceTotal = typeResources.reduce((sum, [resource, data]) => {
                return sum + data.totalCost;
            }, 0);

            // Create HTML content
            const htmlContent = this.generateResourceHtml(proj, sheetName, typeResources, resourceTotal);
            
            // Create and download the file
            const fileName = `${proj.name || 'مشروع'}_${sheetName}_${new Date().toISOString().split('T')[0]}.html`;
            this.downloadHtmlFile(htmlContent, fileName);
            
            console.log(`${sheetName} HTML file downloaded successfully:`, fileName);

        } catch (error) {
            console.error(`Error exporting ${sheetName} to HTML:`, error);
            alert(`حدث خطأ أثناء تصدير ${sheetName} إلى HTML. يرجى المحاولة مرة أخرى.`);
        }
    }

    // Generate HTML content for resource type
    generateResourceHtml(proj, resourceType, typeResources, resourceTotal) {
        const currentDate = new Date().toISOString().split('T')[0];
        const currentTime = new Date().toTimeString().split(' ')[0];
        
        return `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${proj.name} - ${resourceType}</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f5f5f5;
            color: #333;
            line-height: 1.6;
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }
        
        .header .subtitle {
            font-size: 1.2em;
            opacity: 0.9;
        }
        
        .project-info {
            background: #f8f9fa;
            padding: 20px;
            border-bottom: 1px solid #e9ecef;
        }
        
        .project-info table {
            width: 100%;
            border-collapse: collapse;
        }
        
        .project-info td {
            padding: 8px 12px;
            border-bottom: 1px solid #e9ecef;
        }
        
        .project-info td:first-child {
            font-weight: bold;
            color: #495057;
            width: 200px;
        }
        
        .resources-table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        
        .resources-table th {
            background: #495057;
            color: white;
            padding: 15px;
            text-align: center;
            font-weight: 600;
            border: 1px solid #6c757d;
        }
        
        .resources-table td {
            padding: 12px 15px;
            border: 1px solid #dee2e6;
            text-align: center;
        }
        
        .resources-table tr:nth-child(even) {
            background: #f8f9fa;
        }
        
        .resources-table tr:hover {
            background: #e9ecef;
        }
        
        .resource-name {
            text-align: right;
            font-weight: 600;
            color: #495057;
        }
        
        .total-row {
            background: #28a745 !important;
            color: white;
            font-weight: bold;
            font-size: 1.1em;
        }
        
        .total-row td {
            border-color: #28a745;
        }
        
        .usage-details {
            background: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 5px;
            padding: 15px;
            margin: 20px 0;
        }
        
        .usage-details h3 {
            color: #495057;
            margin-bottom: 15px;
            border-bottom: 2px solid #007bff;
            padding-bottom: 5px;
            text-align: center;
        }
        
        .usage-details h4 {
            color: #495057;
            margin-bottom: 10px;
            border-bottom: 1px solid #dee2e6;
            padding-bottom: 5px;
            cursor: pointer;
            user-select: none;
            transition: color 0.2s;
        }
        
        .usage-details h4:hover {
            color: #007bff;
        }
        
        .usage-details h4::after {
            content: ' ▼';
            font-size: 0.8em;
            color: #007bff;
            transition: transform 0.2s;
        }
        
        .usage-details h4.collapsed::after {
            content: ' ▶';
        }
        
        .usage-content {
            overflow: hidden;
            transition: max-height 0.3s ease-in-out;
            max-height: 1000px;
        }
        
        .usage-content.collapsed {
            max-height: 0;
        }
        
        .expand-all-btn {
            background: #007bff;
            color: white;
            border: none;
            border-radius: 5px;
            padding: 8px 16px;
            margin-bottom: 15px;
            cursor: pointer;
            font-size: 0.9em;
            transition: background 0.2s;
        }
        
        .expand-all-btn:hover {
            background: #0056b3;
        }
        
        /* Progress bars and charts */
        .progress-container {
            margin: 20px 0;
            padding: 20px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .progress-item {
            margin-bottom: 15px;
        }
        
        .progress-label {
            display: flex;
            justify-content: space-between;
            margin-bottom: 5px;
            font-weight: 600;
            color: #495057;
        }
        
        .progress-bar {
            width: 100%;
            height: 12px;
            background: #e9ecef;
            border-radius: 6px;
            overflow: hidden;
            position: relative;
        }
        
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #007bff, #0056b3);
            border-radius: 6px;
            transition: width 0.3s ease;
            position: relative;
            box-shadow: inset 0 1px 2px rgba(0,0,0,0.1);
        }
        
        .progress-fill::after {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3));
        }
        
        .progress-fill::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(45deg, transparent 30%, rgba(255,255,255,0.1) 50%, transparent 70%);
            animation: shimmer 2s infinite;
        }
        
        @keyframes shimmer {
            0% { transform: translateX(-100%); }
            100% { transform: translateX(100%); }
        }
        
        /* Bar chart */
        .chart-container {
            margin: 30px 0;
            padding: 25px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .chart-title {
            text-align: center;
            color: #495057;
            margin-bottom: 20px;
            font-size: 1.2em;
            font-weight: 600;
        }
        
        .bar-chart {
            display: flex;
            align-items: end;
            justify-content: space-around;
            height: 200px;
            margin: 20px 0;
            padding: 0 20px;
        }
        
        .bar {
            background: linear-gradient(to top, #007bff, #0056b3);
            border-radius: 4px 4px 0 0;
            position: relative;
            min-width: 40px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
        }
        
        .bar-label {
            position: absolute;
            bottom: -25px;
            left: 50%;
            transform: translateX(-50%);
            font-size: 0.8em;
            color: #6c757d;
            white-space: nowrap;
            text-align: center;
        }
        
        .bar-value {
            position: absolute;
            top: -25px;
            left: 50%;
            transform: translateX(-50%);
            font-size: 0.8em;
            color: #007bff;
            font-weight: 600;
        }
        
        /* Resource type cards */
        .resource-cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }
        
        .resource-card {
            background: white;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-left: 4px solid #007bff;
            transition: transform 0.2s;
        }
        
        .resource-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 16px rgba(0,0,0,0.15);
        }
        
        .resource-card-title {
            color: #007bff;
            font-weight: 600;
            margin-bottom: 10px;
            font-size: 1.1em;
        }
        
        .resource-card-value {
            font-size: 1.5em;
            font-weight: 700;
            color: #495057;
            margin-bottom: 5px;
        }
        
        .resource-card-subtitle {
            color: #6c757d;
            font-size: 0.9em;
        }
        
        /* Icons and visual elements */
        .icon {
            display: inline-block;
            width: 20px;
            height: 20px;
            margin-right: 8px;
            vertical-align: middle;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }
        
        .stat-item {
            background: white;
            padding: 15px;
            border-radius: 6px;
            text-align: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            border-top: 3px solid #007bff;
        }
        
        .stat-number {
            font-size: 1.8em;
            font-weight: 700;
            color: #007bff;
            margin-bottom: 5px;
        }
        
        .stat-label {
            color: #6c757d;
            font-size: 0.9em;
        }
        
        .usage-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
        }
        
        .usage-table th {
            background: #6c757d;
            color: white;
            padding: 10px;
            text-align: center;
            font-size: 0.9em;
        }
        
        .usage-table td {
            padding: 8px 10px;
            border: 1px solid #dee2e6;
            text-align: center;
            font-size: 0.9em;
        }
        
        .footer {
            background: #343a40;
            color: white;
            text-align: center;
            padding: 20px;
            margin-top: 30px;
        }
        
        .footer p {
            margin: 5px 0;
            opacity: 0.8;
        }
        
        @media print {
            body { background: white; }
            .container { box-shadow: none; }
            .header { background: #495057 !important; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 ${resourceType}</h1>
            <div class="subtitle">إدارة الموارد - ${proj.name}</div>
            <div style="margin-top: 15px; opacity: 0.8; font-size: 0.9em;">
                <span style="margin: 0 10px;">📅 ${currentDate}</span>
                <span style="margin: 0 10px;">⏰ ${currentTime}</span>
                <span style="margin: 0 10px;">🏗️ ${proj.type}</span>
            </div>
        </div>
        
        <div class="project-info">
            <table>
                <tr>
                    <td>اسم المشروع:</td>
                    <td>${proj.name}</td>
                    <td>كود المشروع:</td>
                    <td>${proj.code}</td>
                </tr>
                <tr>
                    <td>نوع المشروع:</td>
                    <td>${proj.type}</td>
                    <td>المساحة:</td>
                    <td>${this.formatNumber(proj.area)} م²</td>
                </tr>
                <tr>
                    <td>عدد الأدوار:</td>
                    <td>${proj.floor}</td>
                    <td>تاريخ التصدير:</td>
                    <td>${currentDate}</td>
                </tr>
            </table>
        </div>
        
        <div style="padding: 20px;">
            <!-- Statistics Overview -->
            <div class="stats-grid">
                <div class="stat-item">
                    <div class="stat-number">${typeResources.length}</div>
                    <div class="stat-label">إجمالي الموارد</div>
                </div>
                <div class="stat-item">
                    <div class="stat-number">${this.formatNumber(resourceTotal)}</div>
                    <div class="stat-label">إجمالي التكلفة (جنيه)</div>
                </div>
                <div class="stat-item">
                    <div class="stat-number">${this.formatNumber(typeResources.reduce((sum, [resource, data]) => sum + data.usages.reduce((s, u) => s + u.amount, 0), 0))}</div>
                    <div class="stat-label">إجمالي الكميات</div>
                </div>
                <div class="stat-item">
                    <div class="stat-number">${this.formatNumber(typeResources.reduce((sum, [resource, data]) => sum + data.usages.length, 0))}</div>
                    <div class="stat-label">إجمالي البنود المستخدمة</div>
                </div>
            </div>
            
            <!-- Cost Distribution Progress Bars -->
            <div class="progress-container">
                <h3 style="color: #495057; margin-bottom: 20px; text-align: center;">توزيع التكاليف حسب الموارد</h3>
                ${typeResources.map(([resource, data]) => {
                    const totalCost = data.usages.reduce((sum, u) => sum + u.cost, 0);
                    const percentage = resourceTotal > 0 ? (totalCost / resourceTotal) * 100 : 0;
                    return `
                    <div class="progress-item">
                        <div class="progress-label">
                            <span>${resource}</span>
                            <span>${this.formatNumber(totalCost)} جنيه (${percentage.toFixed(1)}%)</span>
                        </div>
                        <div class="progress-bar">
                            <div class="progress-fill" style="width: ${percentage}%"></div>
                        </div>
                    </div>
                    `;
                }).join('')}
            </div>
            
            <!-- Top Resources Bar Chart -->
            <div class="chart-container">
                <div class="chart-title">أعلى 5 موارد من حيث التكلفة</div>
                <div class="bar-chart">
                    ${typeResources
                        .sort((a, b) => {
                            const costA = a[1].usages.reduce((sum, u) => sum + u.cost, 0);
                            const costB = b[1].usages.reduce((sum, u) => sum + u.cost, 0);
                            return costB - costA;
                        })
                        .slice(0, 5)
                        .map(([resource, data], index) => {
                            const totalCost = data.usages.reduce((sum, u) => sum + u.cost, 0);
                            const maxCost = typeResources.reduce((max, [r, d]) => {
                                const cost = d.usages.reduce((sum, u) => sum + u.cost, 0);
                                return Math.max(max, cost);
                            }, 0);
                            const height = maxCost > 0 ? (totalCost / maxCost) * 100 : 0;
                            return `
                            <div class="bar" style="height: ${height}%">
                                <div class="bar-value">${this.formatNumber(totalCost)}</div>
                                <div class="bar-label">${resource.length > 12 ? resource.substring(0, 12) + '...' : resource}</div>
                            </div>
                            `;
                        }).join('')}
                </div>
            </div>
            
            <!-- Resource Type Distribution Cards -->
            <div class="resource-cards">
                ${typeResources.map(([resource, data]) => {
                    const totalCost = data.usages.reduce((sum, u) => sum + u.cost, 0);
                    const totalQuantity = data.usages.reduce((sum, u) => sum + u.amount, 0);
                    const usageCount = data.usages.length;
                    const percentage = resourceTotal > 0 ? (totalCost / resourceTotal) * 100 : 0;
                    
                    return `
                    <div class="resource-card">
                        <div class="resource-card-title">${resource}</div>
                        <div class="resource-card-value">${this.formatNumber(totalCost)} جنيه</div>
                        <div class="resource-card-subtitle">
                            الكمية: ${this.formatNumber(totalQuantity)} ${data.unit || ''} | 
                            البنود: ${usageCount} | 
                            النسبة: ${percentage.toFixed(1)}%
                        </div>
                    </div>
                    `;
                }).join('')}
            </div>
            
            <table class="resources-table">
                <thead>
                    <tr>
                        <th>اسم المورد</th>
                        <th>الوحدة</th>
                        <th>الكمية الإجمالية</th>
                        <th>سعر الوحدة (جنيه)</th>
                        <th>التكلفة الإجمالية (جنيه)</th>
                    </tr>
                </thead>
                <tbody>
                    ${typeResources.map(([resource, data]) => {
                        const totalCost = data.usages.reduce((sum, u) => sum + u.cost, 0);
                        const totalQuantity = data.usages.reduce((sum, u) => sum + u.amount, 0);
                        const unitPrice = data.totalAmount > 0 ? (data.totalCost / data.totalAmount) : 0;
                        
                        return `
                        <tr>
                            <td class="resource-name">${resource}</td>
                            <td>${data.unit || ''}</td>
                            <td>${this.formatNumber(totalQuantity)}</td>
                            <td>${this.formatNumber(unitPrice)}</td>
                            <td>${this.formatNumber(totalCost)}</td>
                        </tr>
                        `;
                    }).join('')}
                    <tr class="total-row">
                        <td colspan="4">إجمالي تكلفة ${resourceType}</td>
                        <td>${this.formatNumber(resourceTotal)}</td>
                    </tr>
                </tbody>
            </table>
            
            <div class="usage-details">
                <h3>تفاصيل استخدام الموارد</h3>
                <button class="expand-all-btn" onclick="toggleAllUsageDetails()">توسيع/طي الكل</button>
                ${typeResources.map(([resource, data], index) => {
                    const totalCost = data.usages.reduce((sum, u) => sum + u.cost, 0);
                    const totalQuantity = data.usages.reduce((sum, u) => sum + u.amount, 0);
                    
                    return `
                    <div style="margin-bottom: 20px;">
                        <h4 class="usage-header" onclick="toggleUsageDetails(${index})" style="color: #495057; margin-bottom: 10px; border-bottom: 1px solid #dee2e6; padding-bottom: 5px;">
                            ${resource} - إجمالي: ${this.formatNumber(totalCost)} جنيه
                        </h4>
                        <div class="usage-content" id="usage-content-${index}">
                            <table class="usage-table">
                                <thead>
                                    <tr>
                                        <th>رقم البند</th>
                                        <th>اسم البند</th>
                                        <th>الكمية</th>
                                        <th>الوحدة</th>
                                        <th>التكلفة</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${data.usages.map((u, usageIndex) => `
                                    <tr>
                                        <td>${usageIndex + 1}</td>
                                        <td>${u.itemTitle}</td>
                                        <td>${this.formatNumber(u.amount)}</td>
                                        <td>${u.unit}</td>
                                        <td>${this.formatNumber(u.cost)} جنيه</td>
                                    </tr>
                                    `).join('')}
                                </tbody>
                            </table>
                        </div>
                    </div>
                    `;
                }).join('')}
            </div>
        </div>
        
        <div class="footer">
            <p>تم إنشاء هذا التقرير بواسطة نظام حساب تكاليف البناء</p>
            <p>تاريخ التصدير: ${currentDate} | وقت التصدير: ${currentTime}</p>
        </div>
    </div>
    
    <script>
        // Function to toggle individual usage details
        function toggleUsageDetails(index) {
            const content = document.getElementById('usage-content-' + index);
            const header = content.previousElementSibling;
            
            if (content.classList.contains('collapsed')) {
                content.classList.remove('collapsed');
                header.classList.remove('collapsed');
            } else {
                content.classList.add('collapsed');
                header.classList.add('collapsed');
            }
        }
        
        // Function to toggle all usage details
        function toggleAllUsageDetails() {
            const contents = document.querySelectorAll('.usage-content');
            const headers = document.querySelectorAll('.usage-header');
            const button = document.querySelector('.expand-all-btn');
            
            const allCollapsed = Array.from(contents).every(content => 
                content.classList.contains('collapsed')
            );
            
            if (allCollapsed) {
                // Expand all
                contents.forEach(content => content.classList.remove('collapsed'));
                headers.forEach(header => header.classList.remove('collapsed'));
                button.textContent = 'طي الكل';
            } else {
                // Collapse all
                contents.forEach(content => content.classList.add('collapsed'));
                headers.forEach(header => header.classList.add('collapsed'));
                button.textContent = 'توسيع الكل';
            }
        }
        
        // Initialize: start with all details expanded
        document.addEventListener('DOMContentLoaded', function() {
            const button = document.querySelector('.expand-all-btn');
            if (button) {
                button.textContent = 'طي الكل';
            }
        });
    </script>
</body>
</html>`;
    }

    // Download HTML file
    downloadHtmlFile(htmlContent, fileName) {
        const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }

    // Add this function after updateSummaryTotal
    updateResourcesTotals() {
        try {
            // Use the SAME calculation method as renderResourcesAccordion for consistency
            const summary = this.getResourcesSummary();
            
            let materialsTotal = 0;
            let workmanshipTotal = 0;
            let laborTotal = 0;
            
            // Calculate totals using the exact same logic as the display
            Object.entries(summary).forEach(([resource, data]) => {
                try {
                    // Use the same calculation: sum of all usages costs
                    const totalCost = data.usages.reduce((sum, u) => sum + u.cost, 0);
                    
                    if (isNaN(totalCost)) {
                        console.warn(`Invalid cost for resource ${resource}:`, totalCost);
                        return;
                    }
                    
                    switch (data.type) {
                        case 'خامات':
                            materialsTotal += totalCost;
                            break;
                        case 'مصنعيات':
                            workmanshipTotal += totalCost;
                            break;
                        case 'عمالة':
                            laborTotal += totalCost;
                            break;
                        default:
                            console.warn(`Unknown resource type: ${data.type} for resource: ${resource}`);
                    }
                } catch (resourceError) {
                    console.error(`Error processing resource ${resource}:`, resourceError);
                }
            });
            
            // Validate totals
            if (isNaN(materialsTotal) || isNaN(workmanshipTotal) || isNaN(laborTotal)) {
                console.error('Invalid totals calculated:', { materialsTotal, workmanshipTotal, laborTotal });
                this.clearResourcesTotals();
                return;
            }
            
        const grandTotal = materialsTotal + workmanshipTotal + laborTotal;
        
            // Log the calculation for debugging
            console.log('Resources totals calculation:', {
                summary: summary,
                materialsTotal,
                workmanshipTotal,
                laborTotal,
                grandTotal
            });
            
            // Update display elements with formatted numbers
            this.updateResourcesTotalsDisplay(materialsTotal, workmanshipTotal, laborTotal, grandTotal);
            
        } catch (error) {
            console.error('Error updating resources totals:', error);
            this.clearResourcesTotals();
        }
    }
    
    // Helper function to update resources totals display
    updateResourcesTotalsDisplay(materialsTotal, workmanshipTotal, laborTotal, grandTotal) {
        try {
        const materialsEl = document.getElementById('resourcesMaterialsTotal');
        const workmanshipEl = document.getElementById('resourcesWorkmanshipTotal');
        const laborEl = document.getElementById('resourcesLaborTotal');
        const grandEl = document.getElementById('resourcesGrandTotal');
        
            if (materialsEl) materialsEl.textContent = this.formatNumber(materialsTotal) + ' جنيه';
            if (workmanshipEl) workmanshipEl.textContent = this.formatNumber(workmanshipTotal) + ' جنيه';
            if (laborEl) laborEl.textContent = this.formatNumber(laborTotal) + ' جنيه';
            if (grandEl) grandEl.textContent = this.formatNumber(grandTotal) + ' جنيه';
            
            console.log('Resources totals updated successfully:', {
                materials: materialsTotal,
                workmanship: workmanshipTotal,
                labor: laborTotal,
                grand: grandTotal
            });
            
        } catch (error) {
            console.error('Error updating resources totals display:', error);
        }
    }
    
    // Helper function to clear resources totals
    clearResourcesTotals() {
        try {
            const materialsEl = document.getElementById('resourcesMaterialsTotal');
            const workmanshipEl = document.getElementById('resourcesWorkmanshipTotal');
            const laborEl = document.getElementById('resourcesLaborTotal');
            const grandEl = document.getElementById('resourcesGrandTotal');
            
            if (materialsEl) materialsEl.textContent = '0.00 جنيه';
            if (workmanshipEl) workmanshipEl.textContent = '0.00 جنيه';
            if (laborEl) laborEl.textContent = '0.00 جنيه';
            if (grandEl) grandEl.textContent = '0.00 جنيه';
            
            console.log('Resources totals cleared');
            
        } catch (error) {
            console.error('Error clearing resources totals:', error);
        }
    }

    // Helper function to calculate section total
    calculateSectionTotal(sectionBody) {
        try {
        let total = 0;
            const rows = sectionBody.querySelectorAll('.resource-row');
            rows.forEach(row => {
                try {
                    const totalCostEl = row.querySelector('.total-cost .value');
                    if (totalCostEl) {
                        const costText = totalCostEl.textContent;
                        const cost = parseFloat(costText.replace(/[^\d.-]/g, '')) || 0;
                        total += cost;
                    }
                } catch (error) {
                    console.error('Error calculating row cost:', error);
                }
            });
            return this.formatNumber(total) + ' جنيه';
        } catch (error) {
            console.error('Error calculating section total:', error);
            return '0.00 جنيه';
        }
    }

    // Helper function to get section total display
    getSectionTotalDisplay(elementId) {
        try {
            const element = document.getElementById(elementId);
            return element ? element.textContent : '0.00 جنيه';
        } catch (error) {
            console.error('Error getting section total display:', error);
            return '0.00 جنيه';
        }
    }

    // Helper function to calculate summary total
    calculateSummaryTotal() {
        try {
            let total = 0;
            const cards = this.summaryCards.querySelectorAll('.summary-card');
            cards.forEach(card => {
                try {
                    const unitCost = parseFloat(card.cardData?.unitPrice) || 0;
                    const quantity = parseFloat(card.cardData?.quantity) || 0;
                    total += unitCost * quantity;
                } catch (error) {
                    console.error('Error calculating card total:', error);
                }
            });
            return this.formatNumber(total) + ' جنيه';
        } catch (error) {
            console.error('Error calculating summary total:', error);
            return '0.00 جنيه';
        }
    }

    // Helper function to calculate summary selling total
    calculateSummarySellingTotal() {
        try {
            let total = 0;
            const cards = this.summaryCards.querySelectorAll('.summary-card');
            cards.forEach(card => {
                try {
                    const sellPrice = parseFloat(card.dataset.sellPrice) || parseFloat(card.cardData?.unitPrice) || 0;
                    const quantity = parseFloat(card.cardData?.quantity) || 0;
                    total += sellPrice * quantity;
                } catch (error) {
                    console.error('Error calculating card selling total:', error);
                }
            });
            return this.formatNumber(total) + ' جنيه';
        } catch (error) {
            console.error('Error calculating summary selling total:', error);
            return '0.00 جنيه';
        }
    }

    // Helper function to calculate summary final total
    calculateSummaryFinalTotal() {
        try {
            const sellingTotal = parseFloat(this.calculateSummarySellingTotal().replace(/[^\d.-]/g, '')) || 0;
            const supervisionPercent = parseFloat(this.supervisionPercentage?.value) || 0;
            const finalTotal = sellingTotal * (1 + supervisionPercent / 100);
            return this.formatNumber(finalTotal) + ' جنيه';
        } catch (error) {
            console.error('Error calculating summary final total:', error);
            return '0.00 جنيه';
        }
    }

    // Export the summary (البنود) section to a single HTML file
    exportSummaryToHtml() {
        try {
            const proj = this.projects[this.currentProjectId];
            if (!proj) {
                alert('لا يوجد مشروع محدد.');
                return;
            }

            const selectedCards = this.getSelectedCards();
            if (!selectedCards || selectedCards.length === 0) {
                alert('يرجى تحديد بنود للتصدير.');
                return;
            }

            const items = selectedCards.map((card, index) => {
                const d = card.cardData || {};
                const unitPrice = parseFloat(d.unitPrice) || 0;
                const quantity = parseFloat(d.quantity) || 0;
                const wastePercent = parseFloat(d.wastePercent) || 0;
                const operationPercent = parseFloat(d.operationPercent) || 0;
                const riskPercentage = parseFloat(d.riskPercentage) || 0;
                const taxPercentage = parseFloat(d.taxPercentage) || 14;
                const adjustedUnitCost = unitPrice * (1 + wastePercent/100 + operationPercent/100);
                const sellPrice = unitPrice * (1 + riskPercentage/100) * (1 + taxPercentage/100);
                const totalCost = adjustedUnitCost * quantity;
                const totalSell = sellPrice * quantity;
                return {
                    index: index + 1,
                    mainItem: d.mainItem || '',
                    subItem: d.subItem || '',
                    unit: d.unit || '',
                    quantity,
                    adjustedUnitCost,
                    sellPrice,
                    totalCost,
                    totalSell
                };
            });

            const summaryTotal = items.reduce((s, it) => s + it.totalCost, 0);
            const summarySellingTotal = items.reduce((s, it) => s + it.totalSell, 0);
            const supervisionPercent = parseFloat(this.supervisionPercentage?.value) || 0;
            const summaryFinalTotal = summarySellingTotal * (1 + supervisionPercent/100);

            const currentDate = new Date().toISOString().split('T')[0];
            const currentTime = new Date().toTimeString().split(' ')[0];

            const html = this.generateSummaryHtml({ proj, items, summaryTotal, summarySellingTotal, summaryFinalTotal, supervisionPercent, currentDate, currentTime });
            const fileName = `${proj.name || 'مشروع'}_البنود_${currentDate}.html`;
            this.downloadHtmlFile(html, fileName);
        } catch (error) {
            console.error('Error exporting summary to HTML:', error);
            alert('حدث خطأ أثناء التصدير إلى HTML. يرجى المحاولة مرة أخرى.');
        }
    }

    generateSummaryHtml(ctx) {
        const { proj, items, summaryTotal, summarySellingTotal, summaryFinalTotal, supervisionPercent, currentDate, currentTime } = ctx;
        return `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${proj.name} - البنود</title>
    <style>
        * { box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f5f5; color: #333; line-height: 1.6; padding: 20px; }
        .container { max-width: 1200px; margin: 0 auto; background: #fff; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); overflow: hidden; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: #fff; padding: 30px; text-align: center; }
        .header h1 { font-weight: 300; margin: 0 0 8px; }
        .project-info { background: #f8f9fa; padding: 20px; border-bottom: 1px solid #e9ecef; }
        .project-info table { width: 100%; border-collapse: collapse; }
        .project-info td { padding: 8px 12px; border-bottom: 1px solid #e9ecef; }
        .project-info td:first-child { font-weight: 700; color: #495057; width: 200px; }
        .summary-table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        .summary-table th { background: #495057; color: #fff; padding: 12px; text-align: center; border: 1px solid #6c757d; }
        .summary-table td { padding: 10px 12px; border: 1px solid #dee2e6; text-align: center; }
        .totals { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 12px; margin: 20px 0; }
        .total-card { background: #fff; border-left: 4px solid #28a745; border-radius: 8px; padding: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
        .total-card h3 { margin: 0 0 8px; color: #28a745; }
        .total-card .value { font-size: 1.4em; font-weight: 800; color: #495057; }
        .usage-details { background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 6px; padding: 16px; margin: 20px 0; }
        .usage-details h3 { color: #495057; text-align:center; margin-bottom: 10px; border-bottom: 2px solid #007bff; padding-bottom: 5px; }
        .item-card { margin-bottom: 14px; }
        .item-header { color: #495057; margin-bottom: 8px; border-bottom: 1px solid #dee2e6; padding-bottom: 4px; cursor: pointer; user-select: none; }
        .item-header::after { content: ' ▼'; color: #007bff; font-size: 0.85em; }
        .item-header.collapsed::after { content: ' ▶'; }
        .item-content { overflow: hidden; transition: max-height .3s ease; max-height: 800px; }
        .item-content.collapsed { max-height: 0; }
        .footer { background: #343a40; color: #fff; text-align: center; padding: 20px; margin-top: 30px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📑 البنود - ${proj.name}</h1>
            <div style="margin-top:8px;opacity:.85">📅 ${currentDate} • ⏰ ${currentTime}</div>
        </div>
        <div class="project-info">
            <table>
                <tr><td>اسم المشروع:</td><td>${proj.name}</td><td>كود المشروع:</td><td>${proj.code}</td></tr>
                <tr><td>نوع المشروع:</td><td>${proj.type}</td><td>المساحة:</td><td>${this.formatNumber(proj.area)} م²</td></tr>
                <tr><td>عدد الأدوار:</td><td>${proj.floor}</td><td>نسبة الإشراف:</td><td>${supervisionPercent}%</td></tr>
            </table>
        </div>
        <div style="padding:20px;">
            <table class="summary-table">
                <thead>
                    <tr>
                        <th>#</th>
                        <th>البند الرئيسي</th>
                        <th>البند الفرعي</th>
                        <th>الوحدة</th>
                        <th>الكمية</th>
                        <th>تكلفة الوحدة</th>
                        <th>سعر البيع</th>
                        <th>إجمالي التكلفة</th>
                        <th>إجمالي البيع</th>
                    </tr>
                </thead>
                <tbody>
                    ${items.map((it, idx) => `
                        <tr>
                            <td>${idx+1}</td>
                            <td>${it.mainItem}</td>
                            <td>${it.subItem}</td>
                            <td>${it.unit}</td>
                            <td>${this.formatNumber(it.quantity)}</td>
                            <td>${this.formatNumber(it.adjustedUnitCost)}</td>
                            <td>${this.formatNumber(it.sellPrice)}</td>
                            <td>${this.formatNumber(it.totalCost)}</td>
                            <td>${this.formatNumber(it.totalSell)}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            <div class="totals">
                <div class="total-card"><h3>إجمالي التكلفة الأساسية</h3><div class="value">${this.formatNumber(summaryTotal)} جنيه</div></div>
                <div class="total-card"><h3>إجمالي سعر البيع</h3><div class="value">${this.formatNumber(summarySellingTotal)} جنيه</div></div>
                <div class="total-card"><h3>إجمالي سعر البيع النهائي</h3><div class="value">${this.formatNumber(summaryFinalTotal)} جنيه</div></div>
            </div>
            <div class="usage-details">
                <h3>تفاصيل البنود (قابلة للطي)</h3>
                ${items.map((it, idx) => `
                    <div class="item-card">
                        <div class="item-header" onclick="toggleItem(${idx})">${it.mainItem} - ${it.subItem}</div>
                        <div class="item-content" id="item-${idx}">
                            <table class="summary-table">
                                <thead><tr><th>الوصف</th><th>القيمة</th></tr></thead>
                                <tbody>
                                    <tr><td>الكمية</td><td>${this.formatNumber(it.quantity)} ${it.unit}</td></tr>
                                    <tr><td>تكلفة الوحدة بعد الهالك/التشغيل</td><td>${this.formatNumber(it.adjustedUnitCost)} جنيه</td></tr>
                                    <tr><td>سعر البيع للوحدة</td><td>${this.formatNumber(it.sellPrice)} جنيه</td></tr>
                                    <tr><td>إجمالي التكلفة</td><td>${this.formatNumber(it.totalCost)} جنيه</td></tr>
                                    <tr><td>إجمالي البيع</td><td>${this.formatNumber(it.totalSell)} جنيه</td></tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                `).join('')}
            </div>
        </div>
        <div class="footer">
            <p>تم إنشاء هذا التقرير بواسطة نظام حساب تكاليف البناء</p>
        </div>
    </div>
    <script>
        function toggleItem(i){
            const el = document.getElementById('item-'+i);
            const header = el.previousElementSibling;
            el.classList.toggle('collapsed');
            header.classList.toggle('collapsed');
        }
    </script>
</body>
</html>`;
    }

}

// Initialize the calculator when the DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new ConstructionCalculator();
}); 