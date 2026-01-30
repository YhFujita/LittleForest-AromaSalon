window.onload = function () {
    // Initialize LIFF logic if LIFF ID is provided later
    // liff.init({ liffId: "YOUR_LIFF_ID" })...

    // ★ 重要: デプロイされたGASウェブアプリのURLをここに設定してください
    const GAS_API_URL = 'YOUR_GAS_WEB_APP_URL';

    const form = document.getElementById('reservationForm');
    const loading = document.getElementById('loading');
    const success = document.getElementById('successMessage');

    form.addEventListener('submit', function (e) {
        e.preventDefault();

        // Simple Validation
        const menu = document.getElementById('menu').value;
        const datetime = document.getElementById('datetime').value;
        const name = document.getElementById('name').value;
        const phone = document.getElementById('phone').value;
        const notes = document.getElementById('notes').value;

        if (!menu || !datetime || !name || !phone) {
            alert('すべての必須項目を入力してください。');
            return;
        }

        if (GAS_API_URL === 'YOUR_GAS_WEB_APP_URL') {
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
            notes: notes
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
