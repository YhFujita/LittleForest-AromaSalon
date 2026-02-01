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

    const form = document.getElementById('reservationForm');
    const loading = document.getElementById('loading');
    const success = document.getElementById('successMessage');
    const menuSelect = document.getElementById('menu');
    const dateSelect = document.getElementById('datetime');

    // Load Data (Menu + Available Slots)
    fetch(GAS_API_URL + '?action=get_data')
        .then(response => response.json())
        .then(data => {
            if (data.status === 'success') {
                // Populate Menu
                if (data.menu) {
                    menuSelect.innerHTML = '<option value="" disabled selected>メニューを選択してください</option>';

                    let currentSection = null;
                    let currentGroup = null;

                    data.menu.forEach(item => {
                        // Check if section changed (only if new section is provided)
                        const itemSection = item.section || '';

                        if (itemSection) {
                            // If a new section is defined, switch to it
                            if (itemSection !== currentSection) {
                                currentSection = itemSection;
                                currentGroup = document.createElement('optgroup');
                                currentGroup.label = currentSection;
                                menuSelect.appendChild(currentGroup);
                            }
                        }
                        // If itemSection is empty, we keep the currentGroup (inherit)
                        // If no group has been started yet, currentGroup remains null

                        const option = document.createElement('option');
                        option.value = item.id;
                        option.textContent = `${item.name} (${item.duration}分) - ¥${Number(item.price).toLocaleString()}`;

                        if (currentGroup) {
                            currentGroup.appendChild(option);
                        } else {
                            menuSelect.appendChild(option);
                        }
                    });
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

    form.addEventListener('submit', function (e) {
        e.preventDefault();

        // Simple Validation
        const menu = document.getElementById('menu').value;
        const datetime = document.getElementById('datetime').value;
        const name = document.getElementById('name').value;
        const phone = document.getElementById('phone').value;
        const notes = document.getElementById('notes').value;
        const honeypot = document.getElementById('honeypot').value;

        if (!menu || !datetime || !name || !phone) {
            alert('すべての必須項目を入力してください。');
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
            datetime: datetime,
            name: name,
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
