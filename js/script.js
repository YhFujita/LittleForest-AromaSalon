window.onload = function () {
    // Initialize LIFF logic
    // Replace "YOUR_LIFF_ID" with your actual LIFF ID from LINE Developers Console
    // or set it via a global config if preferred. Using a placeholder here for user update.
    liff.init({ liffId: "2009015288-DWfo5Yqy" })
        .then(() => {
            console.log("LIFF Initialized");
            if (!liff.isLoggedIn()) {
                liff.login();
            } else {
                // Auto-fill name from LINE Profile
                liff.getProfile()
                    .then(profile => {
                        const nameInput = document.getElementById('name');
                        if (nameInput && !nameInput.value) {
                            nameInput.value = profile.displayName;
                        }
                    })
                    .catch(err => console.error('Profile fetch failed', err));
            }
        })
        .catch((err) => {
            console.error("LIFF Init failed", err);
        });

    // ★ 重要: デプロイされたGASウェブアプリのURLをここに設定してください
    const GAS_API_URL = 'https://script.google.com/macros/s/AKfycbwZw61TyYkb8fboc8mWRZGZUpzqRWUluykk2cQ4hKXQV83RySPsprsKVL9R8Luy4AbZtw/exec';

    let allMenuItems = [];
    let selectedMainMenuId = null;
    let selectedOptionIds = new Set();

    const form = document.getElementById('reservationForm');
    const loading = document.getElementById('loading');
    const success = document.getElementById('successMessage');
    const menuContainer = document.getElementById('menuContainer');
    const dateSelect = document.getElementById('datetime');

    // Load Data (Menu + Available Slots)
    fetch(GAS_API_URL + '?action=get_data')
        .then(response => response.json())
        .then(data => {
            if (data.status === 'success') {
                // Populate Menu
                if (data.menu) {
                    allMenuItems = data.menu;
                    renderMenuList();
                }
                // Populate Slots
                if (data.slots) {
                    dateSelect.innerHTML = '<option value="" disabled selected>希望日時を選択してください</option>';
                    if (data.slots.length === 0) {
                        const option = document.createElement('option');
                        option.disabled = true;
                        option.textContent = '現在予約できる空き枠がありません';
                        dateSelect.appendChild(option);
                    } else {
                        data.slots.forEach(slot => {
                            const option = document.createElement('option');
                            // slot = { value: "yyyy/MM/dd HH:mm", display: "yyyy年... (Japanese)" }
                            option.value = slot.value;
                            option.textContent = slot.display;
                            dateSelect.appendChild(option);
                        });
                    }
                }
            }
        })
        .catch(err => {
            console.error('Failed to load data', err);
            menuSelect.innerHTML = '<option value="" disabled selected>読み込み失敗</option>';
            dateSelect.innerHTML = '<option value="" disabled selected>読み込み失敗</option>';
        });

    function renderMenuList() {
        menuContainer.innerHTML = '';

        // Add Summary Bar at the top (initially hidden or zeroed)
        const summaryBar = document.createElement('div');
        summaryBar.id = 'menuSummaryBar';
        summaryBar.className = 'menu-summary-bar hidden';
        summaryBar.innerHTML = `
            <span>合計 <span id="totalDuration">0</span>分</span>
            <span class="menu-summary-total">¥<span id="totalPrice">0</span></span>
        `;
        menuContainer.appendChild(summaryBar);

        let currentSection = null;
        let currentSectionDiv = null;

        allMenuItems.forEach(item => {
            const itemSection = item.section || '';

            // Section Header
            if (itemSection && itemSection !== currentSection) {
                currentSection = itemSection;

                const headerDiv = document.createElement('div');
                headerDiv.className = 'menu-section-header';
                headerDiv.innerHTML = `<h3>${itemSection}</h3>`;
                if (item.sectionDesc) {
                    const descP = document.createElement('p');
                    descP.className = 'menu-section-desc';
                    descP.textContent = item.sectionDesc;
                    headerDiv.appendChild(descP);
                }
                menuContainer.appendChild(headerDiv);
            }

            // Menu Card
            const card = document.createElement('div');
            card.className = `menu-item-card ${item.isOption ? 'option-item' : 'main-item'}`;
            if (item.isOption) card.style.opacity = '0.6'; // Dim options initially

            card.innerHTML = `
                <div class="menu-selection-indicator"></div>
                <div class="menu-item-info">
                    <div class="menu-item-main">
                        <span class="menu-item-name">${item.name}</span>
                        ${item.description ? `<span class="menu-item-desc">${item.description}</span>` : ''}
                    </div>
                    <div class="menu-item-meta">
                        <span class="menu-item-price">¥${Number(item.price).toLocaleString()}</span>
                        <span class="menu-item-duration">${item.duration}分</span>
                    </div>
                </div>
            `;

            card.onclick = () => handleMenuSelection(item);
            menuContainer.appendChild(card);
            item._el = card; // Store direct reference
        });
    }

    function handleMenuSelection(item) {
        if (!item.isOption) {
            // Main Menu Selection (Radio-like)
            if (selectedMainMenuId === item.id) {
                selectedMainMenuId = null; // Toggle off
            } else {
                selectedMainMenuId = item.id;
            }
        } else {
            // Option Selection (Checkbox-like)
            if (!selectedMainMenuId) {
                alert('先にメインメニューを選択してください');
                return;
            }
            if (selectedOptionIds.has(item.id)) {
                selectedOptionIds.delete(item.id);
            } else {
                selectedOptionIds.add(item.id);
            }
        }
        updateUIState();
    }

    function updateUIState() {
        allMenuItems.forEach(item => {
            const isSelected = (!item.isOption && selectedMainMenuId === item.id) ||
                (item.isOption && selectedOptionIds.has(item.id));

            item._el.classList.toggle('selected', isSelected);

            // Interaction control for options
            if (item.isOption) {
                item._el.style.opacity = selectedMainMenuId ? '1' : '0.6';
                item._el.style.cursor = selectedMainMenuId ? 'pointer' : 'not-allowed';
            }
        });

        // Update Summary
        const summaryBar = document.getElementById('menuSummaryBar');
        if (selectedMainMenuId) {
            summaryBar.classList.remove('hidden');
            let totalPrice = 0;
            let totalDuration = 0;

            const mainItem = allMenuItems.find(m => m.id === selectedMainMenuId);
            totalPrice += parseInt(mainItem.price);
            totalDuration += parseInt(mainItem.duration);

            selectedOptionIds.forEach(id => {
                const opt = allMenuItems.find(m => m.id === id);
                totalPrice += parseInt(opt.price);
                totalDuration += parseInt(opt.duration);
            });

            document.getElementById('totalPrice').textContent = totalPrice.toLocaleString();
            document.getElementById('totalDuration').textContent = totalDuration;
        } else {
            summaryBar.classList.add('hidden');
            selectedOptionIds.clear(); // Reset options if main is deselected
            updateUIState(); // Recursively sync UI
        }
    }

    form.addEventListener('submit', function (e) {
        e.preventDefault();

        // Simple Validation
        const menu = selectedMainMenuId;
        const datetime = document.getElementById('datetime').value;
        const name = document.getElementById('name').value;
        const phone = document.getElementById('phone').value;
        const notes = document.getElementById('notes').value;
        const honeypot = document.getElementById('honeypot').value;

        if (!menu || !datetime || !name || !phone) {
            alert('メインメニューと日時を選択し、必須項目を入力してください。');
            return;
        }

        // if (GAS_API_URL === '...') check removed because it was flagging the valid URL as invalid.
        if (!GAS_API_URL || GAS_API_URL.includes('YOUR_SCRIPT_ID')) {
            alert('GAS_API_URLが正しく設定されていません。');
            return;
        }

        // UI Updates
        form.classList.add('hidden');
        loading.classList.remove('hidden');

        const data = {
            menu: menu,
            options: Array.from(selectedOptionIds),
            datetime: datetime,
            name: name,
            phone: phone,
            notes: notes,
            honeypot: honeypot, // Bot check
            userId: liff.getContext() ? liff.getContext().userId : null // Get User ID from LIFF Context
        };

        // Call GAS API via fetch
        // Note: Using text/plain to avoid CORS preflight (Simple Request)
        // GAS parses this as postData.contents
        fetch(GAS_API_URL, {
            method: "POST",
            body: JSON.stringify(data),
            headers: {
                "Content-Type": "text/plain;charset=utf-8"
            }
        })
            .then(response => response.json())
            .then(result => {
                if (result.status === 'success') {
                    loading.classList.add('hidden');
                    success.classList.remove('hidden');
                    console.log(result);
                } else {
                    throw new Error(result.message || 'Unknown error');
                }
            })
            .catch(error => {
                loading.classList.add('hidden');
                form.classList.remove('hidden');
                alert('エラーが発生しました: ' + error.message);
                console.error('Fetch error:', error);
            });
    });
};
