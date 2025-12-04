// Initialize Tableau Extension
let dashboard;
let worksheets = [];
window.worksheetColumns = new Map(); // Store columns for each worksheet (global for export)

// Wait for DOM to load
document.addEventListener('DOMContentLoaded', () => {
    // Check if Tableau API is available
    if (typeof tableau === 'undefined') {
        console.error('Tableau Extensions API not loaded!');
        showStatus('Error: Tableau Extensions API not found. Please ensure the extension is loaded from Tableau Desktop.', 'error');
        document.getElementById('worksheetList').innerHTML = 
            '<p style="color: #dc3545;">Error: Tableau Extensions API not found.</p>';
    } else {
        initializeExtension();
    }
});

// Initialize the extension
function initializeExtension() {
    const configureOptions = {
        configure: () => {
            showStatus('Configuration dialog opened', 'info');
            // Simple alert for now
            alert('ZH Export Crosstab Extension\n\nNo configuration needed!\n\nJust select worksheets and click Export.');
        }
    };
    
    tableau.extensions.initializeAsync(configureOptions).then(() => {
        console.log('Extension initialized successfully');
        dashboard = tableau.extensions.dashboardContent.dashboard;
        console.log('Dashboard loaded:', dashboard.name);
        
        // Listen for filter changes on all worksheets
        setupFilterChangeListeners();
        
        loadWorksheets();
    }).catch((error) => {
        console.error('Error initializing extension:', error);
        showStatus('Error initializing extension: ' + error.message, 'error');
        document.getElementById('worksheetList').innerHTML = 
            '<p style="color: #dc3545;">Error initializing extension: ' + error.message + '</p>';
    });
}

// Setup filter change listeners for all worksheets
function setupFilterChangeListeners() {
    try {
        const allWorksheets = dashboard.worksheets;
        console.log('Setting up filter change listeners for', allWorksheets.length, 'worksheets');
        
        allWorksheets.forEach(worksheet => {
            // Listen for filter changes on each worksheet
            worksheet.addEventListener(tableau.TableauEventType.FilterChanged, (event) => {
                console.log('Filter changed on worksheet:', worksheet.name);
                // Clear cached data for this worksheet to force fresh fetch on export
                if (window.worksheetColumns.has(worksheet.name)) {
                    console.log('Clearing cached data for:', worksheet.name);
                    window.worksheetColumns.delete(worksheet.name);
                }
                showStatus('Filter changed - fresh data will be used on export (column selections preserved)', 'info');
                // Don't auto-refresh UI to preserve user's column selections
            });
            
            // Listen for parameter changes as well
            worksheet.addEventListener(tableau.TableauEventType.ParameterChanged, (event) => {
                console.log('Parameter changed on worksheet:', worksheet.name);
                // Clear cached data for this worksheet to force fresh fetch on export
                if (window.worksheetColumns.has(worksheet.name)) {
                    console.log('Clearing cached data for:', worksheet.name);
                    window.worksheetColumns.delete(worksheet.name);
                }
                showStatus('Parameter changed - fresh data will be used on export (column selections preserved)', 'info');
                // Don't auto-refresh UI to preserve user's column selections
            });
        });
        
        // Also listen for dashboard-level parameter changes
        try {
            dashboard.addEventListener(tableau.TableauEventType.ParameterChanged, (event) => {
                console.log('Dashboard parameter changed');
                // Clear all cached data
                window.worksheetColumns.forEach((value, key) => {
                    console.log('Clearing cached data for:', key);
                });
                window.worksheetColumns.clear();
                showStatus('⚠️ Parameter changed - Click 🔄 Refresh button if column structure changed', 'warning');
                // Don't auto-refresh to preserve user's column selections
            });
        } catch (e) {
            console.log('Dashboard parameter listener not available:', e.message);
        }
    } catch (error) {
        console.error('Error setting up filter listeners:', error);
    }
}

// Load all worksheets from the dashboard
async function loadWorksheets() {
    try {
        console.log('Loading worksheets...');
        worksheets = dashboard.worksheets;
        console.log('Found worksheets:', worksheets.length);
        
        const worksheetList = document.getElementById('worksheetList');
        
        if (worksheets.length === 0) {
            worksheetList.innerHTML = '<p style="color: #dc3545;">No worksheets found in this dashboard.</p>';
            return;
        }

        worksheetList.innerHTML = '<p style="color: #666; font-size: 13px;">Loading worksheet information...</p>';
        
        // Fetch column counts for all worksheets
        const worksheetData = [];
        for (const worksheet of worksheets) {
            try {
                const dataTable = await worksheet.getSummaryDataAsync({ maxRows: 1 });
                const filteredColumns = dataTable.columns.filter(col => !col.fieldName.startsWith('AGG('));
                worksheetData.push({
                    worksheet: worksheet,
                    columnCount: filteredColumns.length
                });
                console.log(`Worksheet "${worksheet.name}" has ${filteredColumns.length} columns`);
            } catch (error) {
                console.error(`Error fetching columns for ${worksheet.name}:`, error);
                worksheetData.push({
                    worksheet: worksheet,
                    columnCount: 0
                });
            }
        }
        
        worksheetList.innerHTML = '';
        
        worksheetData.forEach((data, index) => {
            const worksheet = data.worksheet;
            const columnCount = data.columnCount;
            
            console.log('Adding worksheet:', worksheet.name, `(${columnCount} columns)`);
            const div = document.createElement('div');
            div.className = 'worksheet-item';
            
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `worksheet_${index}`;
            checkbox.value = worksheet.name;
            // Only check by default if worksheet has columns
            checkbox.checked = columnCount > 0;
            checkbox.onchange = () => {
                updateExportButton();
                handleWorksheetSelection();
            };
            
            const label = document.createElement('label');
            label.htmlFor = `worksheet_${index}`;
            
            // Add worksheet name with column count
            const nameSpan = document.createElement('span');
            nameSpan.textContent = worksheet.name;
            nameSpan.style.marginRight = '8px';
            
            const countBadge = document.createElement('span');
            countBadge.className = columnCount > 0 ? 'column-badge badge-dimension' : 'column-badge badge-measure';
            countBadge.textContent = `${columnCount} columns`;
            countBadge.style.fontSize = '11px';
            countBadge.style.padding = '2px 6px';
            if (columnCount === 0) {
                countBadge.style.backgroundColor = '#fff3cd';
                countBadge.style.color = '#856404';
                countBadge.style.border = '1px solid #ffc107';
            }
            
            label.appendChild(nameSpan);
            label.appendChild(countBadge);
            
            div.appendChild(checkbox);
            div.appendChild(label);
            worksheetList.appendChild(div);
        });
        
        console.log('Worksheets loaded successfully');
        updateExportButton();
        // Load columns for first worksheet by default
        handleWorksheetSelection();
    } catch (error) {
        console.error('Error loading worksheets:', error);
        document.getElementById('worksheetList').innerHTML = 
            '<p style="color: #dc3545;">Error loading worksheets: ' + error.message + '</p>';
    }
}

// Select all worksheets
function selectAll() {
    const checkboxes = document.querySelectorAll('.worksheet-item input[type="checkbox"]');
    checkboxes.forEach(cb => cb.checked = true);
    updateExportButton();
    handleWorksheetSelection();
}

// Deselect all worksheets
function deselectAll() {
    const checkboxes = document.querySelectorAll('.worksheet-item input[type="checkbox"]');
    checkboxes.forEach(cb => cb.checked = false);
    updateExportButton();
    handleWorksheetSelection();
}

// Update export button state
function updateExportButton() {
    const checkedBoxes = document.querySelectorAll('.worksheet-item input[type="checkbox"]:checked');
    const exportBtn = document.getElementById('exportBtn');
    exportBtn.disabled = checkedBoxes.length === 0;
}

// Handle worksheet selection - load columns
async function handleWorksheetSelection() {
    const checkedBoxes = document.querySelectorAll('.worksheet-item input[type="checkbox"]:checked');
    const selectedWorksheets = Array.from(checkedBoxes).map(cb => cb.value);
    
    const columnContainer = document.getElementById('columnSelectionContainer');
    const columnList = document.getElementById('columnList');
    
    if (selectedWorksheets.length === 0) {
        columnContainer.style.display = 'none';
        return;
    }
    
    // Show column selection
    columnContainer.style.display = 'block';
    columnList.innerHTML = '<p style="color: #666; font-size: 13px;">Loading columns...</p>';
    
    try {
        // If multiple worksheets selected, show all columns grouped by worksheet
        if (selectedWorksheets.length > 1) {
            columnList.innerHTML = '';
            
            // Add info message
            const infoDiv = document.createElement('div');
            infoDiv.className = 'tab-info';
            infoDiv.innerHTML = `<strong>ℹ️ Multiple Worksheets Selected (${selectedWorksheets.length})</strong><br>Use the tabs below to select columns for each worksheet.`;
            columnList.appendChild(infoDiv);

            // Create tabs container
            const tabsContainer = document.createElement('div');
            tabsContainer.className = 'tabs';
            columnList.appendChild(tabsContainer);
            
            for (const worksheetName of selectedWorksheets) {
                const worksheet = worksheets.find(ws => ws.name === worksheetName);
                if (!worksheet) continue;
                
                // Check if we already have columns cached
                if (!window.worksheetColumns.has(worksheetName)) {
                    const dataTable = await worksheet.getSummaryDataAsync({ maxRows: 1 });
                    // Filter out AGG columns
                    const filteredColumns = dataTable.columns.filter(col => !col.fieldName.startsWith('AGG('));
                    window.worksheetColumns.set(worksheetName, filteredColumns);
                }

                const columns = window.worksheetColumns.get(worksheetName);
                
                // Create tab button
                const tabButton = document.createElement('button');
                tabButton.className = 'tab-button' + (selectedWorksheets.indexOf(worksheetName) === 0 ? ' active' : '');
                tabButton.textContent = `${worksheetName} (${columns.length})`;
                tabButton.onclick = (e) => {
                    e.preventDefault();
                    document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
                    tabButton.classList.add('active');
                    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
                    document.getElementById('tab-' + worksheetName).classList.add('active');
                };
                tabsContainer.appendChild(tabButton);
            }

            // Create tab content panels
            for (const worksheetName of selectedWorksheets) {
                const columns = window.worksheetColumns.get(worksheetName);

                // Create tab content
                const tabContent = document.createElement('div');
                tabContent.className = 'tab-content' + (selectedWorksheets.indexOf(worksheetName) === 0 ? ' active' : '');
                tabContent.id = 'tab-' + worksheetName;

                // Add select/deselect buttons for this worksheet
                const buttonGroup = document.createElement('div');
                buttonGroup.style.cssText = 'margin-bottom: 10px; display: flex; gap: 8px;';
                
                const selectAllBtn = document.createElement('button');
                selectAllBtn.textContent = 'Select All';
                selectAllBtn.className = 'btn-select-all';
                selectAllBtn.style.cssText = 'flex: 1; font-size: 13px; padding: 6px 12px;';
                selectAllBtn.onclick = (e) => {
                    e.preventDefault();
                    tabContent.querySelectorAll('.column-item input[type="checkbox"]').forEach(cb => cb.checked = true);
                };
                
                const deselectAllBtn = document.createElement('button');
                deselectAllBtn.textContent = 'Deselect All';
                deselectAllBtn.className = 'btn-deselect-all';
                deselectAllBtn.style.cssText = 'flex: 1; font-size: 13px; padding: 6px 12px;';
                deselectAllBtn.onclick = (e) => {
                    e.preventDefault();
                    tabContent.querySelectorAll('.column-item input[type="checkbox"]').forEach(cb => cb.checked = false);
                };
                
                buttonGroup.appendChild(selectAllBtn);
                buttonGroup.appendChild(deselectAllBtn);
                tabContent.appendChild(buttonGroup);

                // Add columns for this worksheet
                columns.forEach((column, index) => {
                    const div = document.createElement('div');
                    div.className = 'column-item';
                    div.dataset.originalName = column.fieldName;
                    div.dataset.worksheet = worksheetName;

                    // Order number input (editable)
                    const orderInput = document.createElement('input');
                    orderInput.type = 'number';
                    orderInput.className = 'column-order-input';
                    orderInput.value = index + 1;
                    orderInput.min = 1;
                    orderInput.title = 'Change number to reorder';
                    orderInput.style.width = '45px';
                    orderInput.style.textAlign = 'center';
                    
                    // Reorder on change
                    orderInput.addEventListener('change', (e) => {
                        reorderColumns(tabContent, parseInt(e.target.value), div);
                    });

                    const checkbox = document.createElement('input');
                    checkbox.type = 'checkbox';
                    checkbox.id = `column_${worksheetName}_${index}`;
                    checkbox.value = column.fieldName;
                    checkbox.checked = true;
                    checkbox.dataset.index = index;
                    checkbox.dataset.worksheet = worksheetName;

                    // Rename input instead of label
                    const renameInput = document.createElement('input');
                    renameInput.type = 'text';
                    renameInput.className = 'column-rename-input';
                    renameInput.value = column.fieldName;
                    renameInput.title = 'Click to rename for export';
                    renameInput.dataset.originalName = column.fieldName;

                    // Add badge for column type
                    const badge = document.createElement('span');
                    const dataType = column.dataType.toLowerCase();
                    if (dataType.includes('int') || dataType.includes('float') || dataType.includes('real')) {
                        badge.className = 'column-badge badge-measure';
                        badge.textContent = 'Measure';
                    } else {
                        badge.className = 'column-badge badge-dimension';
                        badge.textContent = 'Dimension';
                    }

                    // No drag events needed - using sequence numbers instead

                    div.appendChild(orderInput);
                    div.appendChild(checkbox);
                    div.appendChild(renameInput);
                    div.appendChild(badge);
                    tabContent.appendChild(div);
                });

                columnList.appendChild(tabContent);
            }
        } else {
            // Single worksheet - show columns normally
            const worksheetName = selectedWorksheets[0];
            const worksheet = worksheets.find(ws => ws.name === worksheetName);
            
            if (!worksheet) return;
            
            // Check if we already have columns cached
            if (!window.worksheetColumns.has(worksheetName)) {
                const dataTable = await worksheet.getSummaryDataAsync({ maxRows: 1 });
                // Filter out AGG columns
                const filteredColumns = dataTable.columns.filter(col => !col.fieldName.startsWith('AGG('));
                window.worksheetColumns.set(worksheetName, filteredColumns);
            }

            const columns = window.worksheetColumns.get(worksheetName);
            displayColumnSelection(columns, worksheetName);
        }
        
    } catch (error) {
        console.error('Error loading columns:', error);
        columnList.innerHTML = '<p style="color: #dc3545; font-size: 13px;">Error loading columns</p>';
    }
}

// Display column selection checkboxes
function displayColumnSelection(columns, worksheetName) {
    const columnList = document.getElementById('columnList');
    columnList.innerHTML = '';
    
    if (!columns || columns.length === 0) {
        columnList.innerHTML = '<p style="color: #666; font-size: 13px;">No columns found</p>';
        return;
    }
    
    // Add worksheet name header if provided
    if (worksheetName) {
        const header = document.createElement('div');
        header.style.cssText = 'background: #005eb8; color: white; padding: 10px 15px; border-radius: 6px; margin-bottom: 10px; font-weight: 600; font-size: 14px;';
        header.innerHTML = `📋 Columns from: ${worksheetName}`;
        columnList.appendChild(header);
    }
    
    // Add select/deselect buttons for single worksheet
    const buttonGroup = document.createElement('div');
    buttonGroup.style.cssText = 'margin-bottom: 10px; display: flex; gap: 8px;';
    
    const selectAllBtn = document.createElement('button');
    selectAllBtn.textContent = 'Select All Columns';
    selectAllBtn.className = 'btn-select-all';
    selectAllBtn.style.cssText = 'flex: 1; font-size: 13px; padding: 6px 12px;';
    selectAllBtn.onclick = (e) => {
        e.preventDefault();
        columnList.querySelectorAll('.column-item input[type="checkbox"]').forEach(cb => cb.checked = true);
    };
    
    const deselectAllBtn = document.createElement('button');
    deselectAllBtn.textContent = 'Deselect All Columns';
    deselectAllBtn.className = 'btn-deselect-all';
    deselectAllBtn.style.cssText = 'flex: 1; font-size: 13px; padding: 6px 12px;';
    deselectAllBtn.onclick = (e) => {
        e.preventDefault();
        columnList.querySelectorAll('.column-item input[type="checkbox"]').forEach(cb => cb.checked = false);
    };
    
    buttonGroup.appendChild(selectAllBtn);
    buttonGroup.appendChild(deselectAllBtn);
    columnList.appendChild(buttonGroup);
    
    columns.forEach((column, index) => {
        const div = document.createElement('div');
        div.className = 'column-item';
        div.dataset.originalName = column.fieldName;
        div.dataset.worksheet = worksheetName;

        // Order number input (editable)
        const orderInput = document.createElement('input');
        orderInput.type = 'number';
        orderInput.className = 'column-order-input';
        orderInput.value = index + 1;
        orderInput.min = 1;
        orderInput.title = 'Change number to reorder';
        orderInput.style.width = '45px';
        orderInput.style.textAlign = 'center';

        // Reorder on change
        orderInput.addEventListener('change', (e) => {
            reorderColumns(columnList, parseInt(e.target.value), div);
        });

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = `column_${index}`;
        checkbox.value = column.fieldName;
        checkbox.checked = true;
        checkbox.dataset.index = index;

        // Rename input instead of label
        const renameInput = document.createElement('input');
        renameInput.type = 'text';
        renameInput.className = 'column-rename-input';
        renameInput.value = column.fieldName;
        renameInput.title = 'Click to rename for export';
        renameInput.dataset.originalName = column.fieldName;

        // Add badge for column type
        const badge = document.createElement('span');
        const dataType = column.dataType.toLowerCase();
        if (dataType.includes('int') || dataType.includes('float') || dataType.includes('real')) {
            badge.className = 'column-badge badge-measure';
            badge.textContent = 'Measure';
        } else {
            badge.className = 'column-badge badge-dimension';
            badge.textContent = 'Dimension';
        }

        div.appendChild(orderInput);
        div.appendChild(checkbox);
        div.appendChild(renameInput);
        div.appendChild(badge);
        columnList.appendChild(div);
    });
}

// Note: Select/Deselect all columns functionality is now per-worksheet
// Individual buttons are created dynamically for each worksheet tab

// Refresh worksheets and columns (clear cache and reload)
function refreshAll() {
    console.log('Manual refresh triggered');
    showStatus('Refreshing worksheets and columns...', 'info');
    
    // Clear all cached data
    console.log('Clearing all cached column data...');
    window.worksheetColumns.clear();
    
    // Reload worksheets
    loadWorksheets();
    
    showStatus('✓ Refreshed successfully', 'success');
}

// Collect dashboard filters information
async function collectDashboardFilters() {
    try {
        const filterData = [];
        
        // Add header section
        filterData.push(['DASHBOARD EXPORT SUMMARY']);
        filterData.push(['']);
        filterData.push(['Dashboard Name:', dashboard.name]);
        filterData.push(['Export Date:', new Date().toLocaleString()]);
        filterData.push(['Export User:', tableau.extensions.environment.mode || 'Desktop']);
        filterData.push(['']);
        filterData.push(['ACTIVE FILTERS']);
        filterData.push(['Filter Name', 'Applied Values']);
        filterData.push(['']);
        
        let hasFilters = false;
        
        // Get filters from all worksheets
        for (const ws of worksheets) {
            try {
                const filters = await ws.getFiltersAsync();
                
                if (filters.length > 0) {
                    // Add worksheet section header
                    filterData.push([`Worksheet: ${ws.name}`, '']);
                    
                    for (const filter of filters) {
                        hasFilters = true;
                        let filterValues = '';
                        
                        // Handle different filter types
                        if (filter.filterType === 'categorical') {
                            if (filter.appliedValues && filter.appliedValues.length > 0) {
                                const values = filter.appliedValues.map(v => v.value);
                                filterValues = values.slice(0, 100).join(', ');
                                if (values.length > 100) {
                                    filterValues += ` ... (${values.length - 100} more)`;
                                }
                                // Truncate if still too long (Excel limit is ~32K)
                                if (filterValues.length > 30000) {
                                    filterValues = filterValues.substring(0, 30000) + '... (truncated)';
                                }
                            } else if (filter.isExcludeMode) {
                                const excluded = (filter.excludedValues || []).map(v => v.value);
                                filterValues = 'Excluded: ' + excluded.slice(0, 100).join(', ');
                                if (excluded.length > 100) {
                                    filterValues += ` ... (${excluded.length - 100} more)`;
                                }
                                if (filterValues.length > 30000) {
                                    filterValues = filterValues.substring(0, 30000) + '... (truncated)';
                                }
                            }
                        } else if (filter.filterType === 'range') {
                            if (filter.minValue !== null && filter.maxValue !== null) {
                                filterValues = `${filter.minValue} to ${filter.maxValue}`;
                            } else if (filter.minValue !== null) {
                                filterValues = `>= ${filter.minValue}`;
                            } else if (filter.maxValue !== null) {
                                filterValues = `<= ${filter.maxValue}`;
                            }
                        } else if (filter.filterType === 'relative-date') {
                            filterValues = filter.periodType || 'Relative Date Filter';
                        }
                        
                        filterData.push([`  ${filter.fieldName}`, filterValues || 'All']);
                    }
                    
                    filterData.push(['']); // Empty row between worksheets
                }
            } catch (error) {
                console.log(`Could not get filters for ${ws.name}:`, error.message);
            }
        }
        
        // Get dashboard-level parameters
        try {
            const parameters = await dashboard.getParametersAsync();
            if (parameters.length > 0) {
                filterData.push(['DASHBOARD PARAMETERS', '']);
                for (const param of parameters) {
                    filterData.push([`  ${param.name}`, param.currentValue.value || 'Not Set']);
                }
                hasFilters = true;
            }
        } catch (error) {
            console.log('Could not get parameters:', error.message);
        }
        
        if (!hasFilters) {
            filterData.push(['No filters currently applied', '']);
        }
        
        return filterData;
        
    } catch (error) {
        console.error('Error collecting dashboard filters:', error);
        return [
            ['DASHBOARD EXPORT SUMMARY'],
            [''],
            ['Dashboard Name:', dashboard.name],
            ['Export Date:', new Date().toLocaleString()],
            [''],
            ['Error collecting filters:', error.message]
        ];
    }
}

// Show status message
function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
    
    if (type === 'success' || type === 'error') {
        setTimeout(() => {
            statusDiv.className = 'status';
        }, 5000);
    }
}

// Show/hide loading indicator
function setLoading(isLoading) {
    const loading = document.getElementById('loading');
    const exportBtn = document.getElementById('exportBtn');
    
    if (isLoading) {
        loading.classList.add('active');
        exportBtn.disabled = true;
    } else {
        loading.classList.remove('active');
        updateExportButton();
    }
}



// Convert crosstab data to proper structure
function processCrosstabData(data) {
    if (!data || data.length === 0) return [];
    
    // The data comes in a format where:
    // - Each row contains all the information
    // - We need to identify dimension vs measure columns
    const processedData = [];
    
    // Simply return the data as-is for crosstab - Excel will handle the structure
    // The key is that we're using getCrosstabDataAsync which preserves the pivot structure
    return data;
}

// Get the selected export format
function getExportFormat() {
    const formatRadio = document.querySelector('input[name="exportFormat"]:checked');
    const format = formatRadio ? formatRadio.value : 'crosstab';
    console.log('Selected export format:', format);
    return format;
}

// Reorder columns based on sequence number
function reorderColumns(tabContent, newPosition, movedItem) {
    const allItems = [...tabContent.querySelectorAll('.column-item')];
    const currentIndex = allItems.indexOf(movedItem);
    
    // Validate position
    if (newPosition < 1) newPosition = 1;
    if (newPosition > allItems.length) newPosition = allItems.length;
    
    const newIndex = newPosition - 1;
    if (newIndex === currentIndex) return;
    
    // Remove item from current position
    movedItem.remove();
    
    // Insert at new position
    if (newIndex >= allItems.length - 1) {
        tabContent.appendChild(movedItem);
    } else {
        const referenceItem = allItems[newIndex];
        tabContent.insertBefore(movedItem, referenceItem);
    }
    
    // Update all sequence numbers
    updateColumnOrder(tabContent);
}

function updateColumnOrder(tabContent) {
    const items = tabContent.querySelectorAll('.column-item');
    items.forEach((item, index) => {
        const orderInput = item.querySelector('.column-order-input');
        if (orderInput) {
            orderInput.value = index + 1;
        }
        const checkbox = item.querySelector('input[type="checkbox"]');
        if (checkbox) {
            checkbox.dataset.index = index;
        }
    });
}

// Export selected worksheets to Excel
async function exportToExcel() {
    const checkedBoxes = document.querySelectorAll('.worksheet-item input[type="checkbox"]:checked');
    const selectedWorksheets = Array.from(checkedBoxes).map(cb => cb.value);

    if (selectedWorksheets.length === 0) {
        showStatus('Please select at least one worksheet to export.', 'error');
        return;
    }

    // Get selected columns grouped by worksheet in display order
    const worksheetColumns = new Map();

    // Handle multiple worksheets (with tabs)
    document.querySelectorAll('.tab-content').forEach(tabContent => {
        const worksheetName = tabContent.id.replace('tab-', '');
        if (!selectedWorksheets.includes(worksheetName)) return;

        const columnItems = tabContent.querySelectorAll('.column-item');
        const indices = [];
        const names = [];
        const originalNames = [];

        columnItems.forEach(item => {
            const checkbox = item.querySelector('input[type="checkbox"]');
            if (checkbox && checkbox.checked) {
                const renameInput = item.querySelector('.column-rename-input');
                const originalName = renameInput ? renameInput.dataset.originalName : checkbox.value;
                const newName = renameInput ? (renameInput.value.trim() || originalName) : originalName;

                // Find original index from cached columns
                const cachedColumns = window.worksheetColumns?.get(worksheetName) || [];
                const originalIndex = cachedColumns.findIndex(col => col.fieldName === originalName);

                if (originalIndex >= 0) {
                    indices.push(originalIndex);
                    names.push(newName);
                    originalNames.push(originalName);
                }
            }
        });

        // Always add to map, even if no columns selected (to distinguish from "not configured")
        worksheetColumns.set(worksheetName, { indices, names, originalNames });
    });
    
    // Handle single worksheet (no tabs, direct column list)
    if (selectedWorksheets.length === 1 && worksheetColumns.size === 0) {
        const worksheetName = selectedWorksheets[0];
        const columnList = document.getElementById('columnList');
        const columnItems = columnList.querySelectorAll('.column-item');
        const indices = [];
        const names = [];
        const originalNames = [];

        columnItems.forEach(item => {
            const checkbox = item.querySelector('input[type="checkbox"]');
            if (checkbox && checkbox.checked) {
                const renameInput = item.querySelector('.column-rename-input');
                const originalName = renameInput ? renameInput.dataset.originalName : checkbox.value;
                const newName = renameInput ? (renameInput.value.trim() || originalName) : originalName;

                // Find original index from cached columns
                const cachedColumns = window.worksheetColumns?.get(worksheetName) || [];
                const originalIndex = cachedColumns.findIndex(col => col.fieldName === originalName);

                if (originalIndex >= 0) {
                    indices.push(originalIndex);
                    names.push(newName);
                    originalNames.push(originalName);
                }
            }
        });

        // Always add to map, even if no columns selected (to distinguish from "not configured")
        worksheetColumns.set(worksheetName, { indices, names, originalNames });
        if (indices.length === 0) {
            console.log(`Single worksheet mode: No columns selected for ${worksheetName}, will skip this worksheet`);
        } else {
            console.log(`Single worksheet mode: ${names.length} columns selected for ${worksheetName}`, { indices, names });
        }
    }    // Save column selection config temporarily BEFORE clearing cache
    console.log('Saving user column selections before export...');
    const columnSelectionConfig = new Map(worksheetColumns);
    
    // Log what we saved
    columnSelectionConfig.forEach((config, wsName) => {
        console.log(`Saved selections for ${wsName}:`, config.names);
    });
    
    // Clear all cached worksheet column data to force fresh fetch with current filters
    console.log('Clearing cached worksheet data to fetch fresh filtered data...');
    const keysToDelete = [];
    window.worksheetColumns.forEach((value, key) => {
        if (selectedWorksheets.includes(key)) {
            keysToDelete.push(key);
        }
    });
    keysToDelete.forEach(key => {
        console.log('Clearing cached data for:', key);
        window.worksheetColumns.delete(key);
    });
    
    // Get aggregation option (default is distinct/aggregated, checkbox enables duplicates)
    const includeDuplicateRows = document.getElementById('includeDuplicateRows')?.checked || false;
    const showDistinctOnly = !includeDuplicateRows; // Invert: default is distinct, checkbox adds duplicates
    const includeDashboardFilters = document.getElementById('includeDashboardFilters')?.checked || false;

    console.log('Export options:', {
        worksheets: selectedWorksheets.length,
        worksheetColumns: Array.from(worksheetColumns.entries()).map(([name, cols]) => `${name}: ${cols.names.length} columns`),
        distinctOnly: showDistinctOnly,
        includeDuplicates: includeDuplicateRows,
        includeFilters: includeDashboardFilters
    });

    setLoading(true);
    showStatus('Exporting worksheets...', 'info');

    try {
        const workbook = XLSX.utils.book_new();

        // Add Dashboard Filters summary sheet if requested
        if (includeDashboardFilters) {
            showStatus('Collecting dashboard filters...', 'info');
            const filterSummary = await collectDashboardFilters();
            if (filterSummary && filterSummary.length > 0) {
                const filterSheet = XLSX.utils.aoa_to_sheet(filterSummary);
                
                // Set column widths
                filterSheet['!cols'] = [
                    { wch: 30 },  // Filter Name
                    { wch: 50 }   // Values
                ];
                
                XLSX.utils.book_append_sheet(workbook, filterSheet, 'Dashboard Filters');
                console.log('Added Dashboard Filters sheet');
            }
        }

        for (const worksheetName of selectedWorksheets) {
            const worksheet = worksheets.find(ws => ws.name === worksheetName);

            if (!worksheet) continue;

            showStatus(`Processing: ${worksheetName}...`, 'info');

            try {
                // Force fetch fresh data respecting current dashboard filters
                console.log(`Fetching fresh data for ${worksheetName} with current filters applied...`);
                const dataTable = await worksheet.getSummaryDataAsync();
                console.log(`Retrieved ${dataTable.data.length} rows for ${worksheetName}`);
                
                // Filter out AGG columns from the fresh data
                const filteredColumnIndices = [];
                const filteredColumnNames = [];
                dataTable.columns.forEach((col, idx) => {
                    if (!col.fieldName.startsWith('AGG(')) {
                        filteredColumnIndices.push(idx);
                        filteredColumnNames.push(col.fieldName);
                    }
                });
                console.log(`Filtered out AGG columns: ${dataTable.columns.length} -> ${filteredColumnIndices.length} columns`);

                let data;

                // Get columns specific to this worksheet from saved config
                const wsColumns = columnSelectionConfig.get(worksheetName);

                if (wsColumns && wsColumns.originalNames && wsColumns.originalNames.length > 0) {
                    // User has selected specific columns - match by field name (not index)
                    console.log(`Attempting to match ${wsColumns.originalNames.length} selected columns by name for ${worksheetName}`);
                    const mappedIndices = [];
                    const mappedNames = [];
                    
                    // Match selected columns by original field name in fresh data
                    wsColumns.originalNames.forEach((originalName, idx) => {
                        // Find this column in the fresh dataTable by field name
                        const freshColIndex = dataTable.columns.findIndex(col => col.fieldName === originalName);
                        
                        if (freshColIndex >= 0 && !dataTable.columns[freshColIndex].fieldName.startsWith('AGG(')) {
                            // Column exists in fresh data and is not AGG
                            mappedIndices.push(freshColIndex);
                            mappedNames.push(wsColumns.names[idx]);
                            console.log(`  ✓ Matched "${originalName}" → index ${freshColIndex}`);
                        } else {
                            console.log(`  ✗ Column "${originalName}" not found in fresh data or is AGG`);
                        }
                    });
                    
                    if (mappedIndices.length > 0) {
                        console.log(`Exporting ${mappedIndices.length} matched columns for ${worksheetName}`, { mappedNames });
                        data = filterColumns(dataTable, mappedIndices, mappedNames, showDistinctOnly);
                    } else {
                        console.log('No columns matched in fresh data, skipping worksheet (no valid columns selected)');
                        data = null; // Skip this worksheet
                    }
                } else if (wsColumns && wsColumns.originalNames && wsColumns.originalNames.length === 0) {
                    // User explicitly deselected all columns - skip this worksheet
                    console.log(`No columns selected for ${worksheetName}, skipping worksheet export`);
                    data = null;
                } else {
                    // Column selection UI wasn't used or worksheet wasn't in config - export all non-AGG columns
                    console.log(`No column selection config for ${worksheetName}, exporting all non-AGG columns (distinct: ${showDistinctOnly})`);
                    data = filterColumns(dataTable, filteredColumnIndices, filteredColumnNames, showDistinctOnly);
                }                if (data && data.length > 0) {
                    const sheetName = sanitizeSheetName(worksheetName);
                    const ws = XLSX.utils.aoa_to_sheet(data);

                    const colWidths = calculateColumnWidths(data);
                    ws['!cols'] = colWidths;

                    XLSX.utils.book_append_sheet(workbook, ws, sheetName);
                    console.log(`✓ Added worksheet "${sheetName}" to export`);
                } else if (data === null) {
                    console.log(`⊗ Skipped worksheet "${worksheetName}" - no columns selected`);
                } else {
                    console.log(`⚠ Worksheet "${worksheetName}" has no data to export`);
                }
            } catch (error) {
                console.error('Error processing worksheet', worksheetName, ':', error);
                showStatus(`Error processing ${worksheetName}: ${error.message}`, 'error');
            }
        }

        // Generate Excel file
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
        const filename = `Tableau_Export_${timestamp}.xlsx`;

        XLSX.writeFile(workbook, filename);

        const aggregationMsg = showDistinctOnly ? ' (aggregated - unique rows only)' : ' (includes all duplicate rows)';
        showStatus(`✓ Successfully exported ${selectedWorksheets.length} worksheet(s) to ${filename}${aggregationMsg}`, 'success');
    } catch (error) {
        console.error('Export error:', error);
        showStatus(`✗ Error exporting: ${error.message}`, 'error');
    } finally {
        setLoading(false);
    }
}// Filter columns and optionally get distinct values
function filterColumns(dataTable, selectedIndices, selectedNames, distinctOnly) {
    const data = [];
    
    // Add headers (selected columns only)
    data.push(selectedNames);
    
    if (distinctOnly) {
        // Get distinct values for selected columns
        const distinctRows = new Set();
        
        for (let i = 0; i < dataTable.data.length; i++) {
            const row = [];
            for (const colIndex of selectedIndices) {
                row.push(dataTable.data[i][colIndex].formattedValue);
            }
            // Create unique key from row values
            const rowKey = row.join('|||');
            if (!distinctRows.has(rowKey)) {
                distinctRows.add(rowKey);
                data.push(row);
            }
        }
        
        console.log(`Filtered to ${data.length - 1} distinct rows from ${dataTable.data.length} total rows`);
    } else {
        // Export all rows with selected columns
        for (let i = 0; i < dataTable.data.length; i++) {
            const row = [];
            for (const colIndex of selectedIndices) {
                row.push(dataTable.data[i][colIndex].formattedValue);
            }
            data.push(row);
        }
    }
    
    return data;
}

// Get distinct values from all columns
function getDistinctValues(dataTable, selectedIndices) {
    const data = [];
    const columns = dataTable.columns;
    const indices = selectedIndices || columns.map((_, i) => i);
    
    // Add headers
    const headers = indices.map(i => columns[i].fieldName);
    data.push(headers);
    
    // Track distinct rows
    const distinctRows = new Set();
    
    for (let i = 0; i < dataTable.data.length; i++) {
        const row = [];
        for (const colIndex of indices) {
            row.push(dataTable.data[i][colIndex].formattedValue);
        }
        
        // Create unique key
        const rowKey = row.join('|||');
        if (!distinctRows.has(rowKey)) {
            distinctRows.add(rowKey);
            data.push(row);
        }
    }
    
    console.log(`Reduced from ${dataTable.data.length} rows to ${data.length - 1} distinct rows`);
    return data;
}

// Convert summary data to crosstab/pivot format
function convertToCrosstabFormat(dataTable) {
    try {
        const columns = dataTable.columns;
        const data = dataTable.data;
        
        console.log('Converting to crosstab, data rows:', data.length);
        console.log('Columns:', columns.map(c => `${c.fieldName} (${c.dataType})`).join(', '));
        
        if (data.length === 0) {
            return [columns.map(col => col.fieldName)];
        }
        
        // Identify column types - be more flexible with detection
        const dimensions = [];
        const measures = [];
        
        columns.forEach((col, idx) => {
            const dataType = col.dataType.toLowerCase();
            
            // Check if it's a measure (numeric)
            if (dataType.includes('float') || dataType.includes('int') || 
                dataType === 'real' || dataType === 'number') {
                measures.push({ index: idx, name: col.fieldName });
            } else {
                // Everything else is a dimension
                dimensions.push({ index: idx, name: col.fieldName });
            }
        });
        
        console.log('Detected - Dimensions:', dimensions.map(d => d.name).join(', '));
        console.log('Detected - Measures:', measures.map(m => m.name).join(', '));
        
        // Try to create pivot if we have dimensions and measures
        if (dimensions.length >= 2 && measures.length >= 1) {
            console.log('Creating pivot with', dimensions.length, 'dimensions');
            return createSimplePivot(data, dimensions, measures);
        } else if (dimensions.length === 1 && measures.length >= 1) {
            console.log('Creating simple grouped table (1 dimension)');
            return createGroupedTable(data, dimensions, measures);
        } else if (dimensions.length >= 1 && measures.length >= 1) {
            console.log('Creating grouped table (multiple dimensions)');
            return createGroupedTable(data, dimensions, measures);
        } else {
            // No clear structure, return as-is
            console.log('No pivot structure detected, returning raw data');
            return convertTableauDataToArray(dataTable);
        }
    } catch (error) {
        console.error('Error in convertToCrosstabFormat:', error);
        return convertTableauDataToArray(dataTable);
    }
}

// Simple 2D pivot - rows x columns
function createSimplePivot(data, dimensions, measures) {
    // Use first 2 dimensions for pivot (even if there are more)
    const rowDim = dimensions[0];
    const colDim = dimensions[1];
    
    console.log(`Creating pivot: Rows="${rowDim.name}" x Columns="${colDim.name}"`);
    
    // Get unique values
    const rowValues = [...new Set(data.map(row => row[rowDim.index].formattedValue))];
    const colValues = [...new Set(data.map(row => row[colDim.index].formattedValue))];
    
    // Sort if they look like numbers or dates
    try {
        if (!isNaN(rowValues[0])) rowValues.sort((a, b) => Number(a) - Number(b));
        else rowValues.sort();
        
        if (!isNaN(colValues[0])) colValues.sort((a, b) => Number(a) - Number(b));
        else colValues.sort();
    } catch (e) {
        rowValues.sort();
        colValues.sort();
    }
    
    console.log('Row values:', rowValues.join(', '));
    console.log('Column values:', colValues.join(', '));
    console.log('Measures:', measures.map(m => m.name).join(', '));
    
    const result = [];
    
    // Build header structure
    if (measures.length === 1) {
        // Single measure - simpler header
        const header = [rowDim.name];
        colValues.forEach(col => header.push(col));
        result.push(header);
        
        // Second row with measure name
        const measureHeader = [measures[0].name];
        colValues.forEach(() => measureHeader.push(''));
        result.push(measureHeader);
        
    } else {
        // Multiple measures - hierarchical header
        const header1 = [''];
        colValues.forEach(col => {
            header1.push(col);
            for (let i = 1; i < measures.length; i++) {
                header1.push('');
            }
        });
        result.push(header1);
        
        const header2 = [rowDim.name];
        colValues.forEach(() => {
            measures.forEach(m => header2.push(m.name));
        });
        result.push(header2);
    }
    
    // Build data rows
    rowValues.forEach(rowVal => {
        const dataRow = [rowVal];
        
        colValues.forEach(colVal => {
            measures.forEach(measure => {
                // Find matching data point
                const match = data.find(d => 
                    d[rowDim.index].formattedValue === rowVal && 
                    d[colDim.index].formattedValue === colVal
                );
                dataRow.push(match ? match[measure.index].formattedValue : '');
            });
        });
        
        result.push(dataRow);
    });
    
    console.log(`Pivot result: ${result.length} rows x ${result[0].length} columns`);
    
    return result;
}

// Create simple grouped table (1 dimension + measures)
function createGroupedTable(data, dimensions, measures) {
    const result = [];
    const headerRow = dimensions.map(d => d.name).concat(measures.map(m => m.name));
    result.push(headerRow);
    
    const grouped = new Map();
    
    data.forEach(row => {
        const key = dimensions.map(d => row[d.index].formattedValue).join('|||');
        
        if (!grouped.has(key)) {
            const dataRow = dimensions.map(d => row[d.index].formattedValue);
            measures.forEach(m => dataRow.push(row[m.index].formattedValue));
            grouped.set(key, dataRow);
        }
    });
    
    grouped.forEach(row => result.push(row));
    return result;
}



// Convert Tableau DataTable to array format
function convertTableauDataToArray(dataTable) {
    const data = [];
    
    // Add headers
    const headers = dataTable.columns.map(col => col.fieldName);
    data.push(headers);
    
    // Add data rows
    for (let i = 0; i < dataTable.data.length; i++) {
        const row = [];
        for (let j = 0; j < dataTable.columns.length; j++) {
            const value = dataTable.data[i][j];
            row.push(value.formattedValue);
        }
        data.push(row);
    }
    
    return data;
}

// Pivot table data - simple and fast
function pivotTableData(dataTable) {
    try {
        const columns = dataTable.columns;
        const rows = dataTable.data;
        
        if (!rows || rows.length === 0) {
            return [columns.map(col => col.fieldName)];
        }
        
        // Find dimensions (text/date) and measures (numbers)
        const dims = [];
        const measures = [];
        
        for (let i = 0; i < columns.length; i++) {
            const dtype = (columns[i].dataType || '').toLowerCase();
            if (dtype.includes('int') || dtype.includes('float') || dtype.includes('real')) {
                measures.push({ idx: i, name: columns[i].fieldName });
            } else {
                dims.push({ idx: i, name: columns[i].fieldName });
            }
        }
        
        // Need at least 2 dimensions and 1 measure for pivot
        if (dims.length < 2 || measures.length < 1 || rows.length > 5000) {
            return convertTableauDataToArray(dataTable);
        }
        
        // Use last dimension for columns, rest for rows
        const colDim = dims[dims.length - 1];
        const rowDims = dims.slice(0, -1);
        
        // Get unique values
        const colVals = [];
        const seen = new Set();
        for (let i = 0; i < rows.length && colVals.length < 100; i++) {
            const val = rows[i][colDim.idx].formattedValue;
            if (!seen.has(val)) {
                colVals.push(val);
                seen.add(val);
            }
        }
        
        // Build result array
        const result = [];
        
        // Header
        const header = [...rowDims.map(d => d.name)];
        for (let cv of colVals) {
            for (let m of measures) {
                header.push(cv);
            }
        }
        result.push(header);
        
        // Get unique row keys
        const rowKeys = new Map();
        for (let row of rows) {
            const key = rowDims.map(d => row[d.idx].formattedValue).join('|||');
            if (!rowKeys.has(key)) {
                rowKeys.set(key, rowDims.map(d => row[d.idx].formattedValue));
            }
        }
        
        // Build data rows
        for (let [key, rowVals] of rowKeys) {
            const dataRow = [...rowVals];
            
            for (let colVal of colVals) {
                for (let measure of measures) {
                    let found = false;
                    for (let row of rows) {
                        let match = true;
                        for (let i = 0; i < rowDims.length; i++) {
                            if (row[rowDims[i].idx].formattedValue !== rowVals[i]) {
                                match = false;
                                break;
                            }
                        }
                        if (match && row[colDim.idx].formattedValue === colVal) {
                            dataRow.push(row[measure.idx].formattedValue);
                            found = true;
                            break;
                        }
                    }
                    if (!found) dataRow.push('');
                }
            }
            
            result.push(dataRow);
        }
        
        return result;
        
    } catch (e) {
        // Any error, just return simple table
        return convertTableauDataToArray(dataTable);
    }
}

// Sanitize sheet name for Excel compatibility
function sanitizeSheetName(name) {
    // Excel sheet names can't exceed 31 characters and can't contain: \ / ? * [ ]
    let sanitized = name.replace(/[\\\/\?\*\[\]]/g, '_');
    if (sanitized.length > 31) {
        sanitized = sanitized.substring(0, 31);
    }
    return sanitized;
}

// Calculate column widths for better formatting
function calculateColumnWidths(data) {
    if (!data || data.length === 0) return [];
    
    const colWidths = [];
    const maxCols = Math.max(...data.map(row => row.length));
    
    for (let col = 0; col < maxCols; col++) {
        let maxWidth = 10; // Minimum width
        
        for (const row of data) {
            if (row[col]) {
                const cellLength = String(row[col]).length;
                maxWidth = Math.max(maxWidth, Math.min(cellLength, 50)); // Cap at 50
            }
        }
        
        colWidths.push({ wch: maxWidth });
    }
    
    return colWidths;
}
