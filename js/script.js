window.onload = function () {
    // Initialize LIFF logic if LIFF ID is provided later
    // liff.init({ liffId: "YOUR_LIFF_ID" })...

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
                    data.menu.forEach(item => {
                        const option = document.createElement('option');
                        option.value = item.id;
                        option.textContent = `${item.name} (${item.duration}分) - ¥${Number(item.price).toLocaleString()}`;
                        menuSelect.appendChild(option);
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
                            option.value = slot; // "yyyy/MM/dd HH:mm"
                            option.textContent = slot;
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

        if (GAS_API_URL === 'https://script.google.com/macros/s/AKfycbwZw61TyYkb8fboc8mWRZGZUpzqRWUluykk2cQ4hKXQV83RySPsprsKVL9R8Luy4AbZtw/exec') {
            alert('GAS_API_URLが設定されていません。script.jsを確認してください。');
            return;
        }

        // UI Updates
        form.classList.add('hidden');
        loading.classList.remove('hidden');

        const data = {
            menu: menu,
            datetime: datetime,
            name: name,
            phone: phone,
            notes: notes,
            honeypot: honeypot // Bot check
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
