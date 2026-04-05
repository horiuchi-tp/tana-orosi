const GAS_API_URL = httpsscript.google.commacrossAKfycbyo1t1sX5GyrCx22yfJOFNUJv6CesapJ7xGoFk947IDFF01glOPJLU5S3X3bizQE3tYBwexec;

let html5QrCode = null;
let isScanning = false;
let scannedDataMap = new Map(); 
let toastTimeout;
let wakeLock = null; 

document.addEventListener(DOMContentLoaded, () = {
    loadFromLocalStorage();
    updateDataCount();
    setupEventListeners();
});

function setupEventListeners() {
    document.getElementById('btnStart').addEventListener('click', startScanner);
    document.getElementById('btnStop').addEventListener('click', stopScanner);
    document.getElementById('btnSend').addEventListener('click', sendDataToGAS);
}

function saveToLocalStorage() {
    localStorage.setItem('inventoryDataMap', JSON.stringify(Array.from(scannedDataMap.entries())));
}

function loadFromLocalStorage() {
    const savedData = localStorage.getItem('inventoryDataMap');
    if (savedData) {
        scannedDataMap = new Map(JSON.parse(savedData));
    }
}

function clearLocalStorage() {
    localStorage.removeItem('inventoryDataMap');
    scannedDataMap.clear();
    updateDataCount();
}

function updateDataCount() {
    document.getElementById('dataCount').innerText = scannedDataMap.size;
}

function showToast(message, qty) {
    const toast = document.getElementById('toast');
    toast.innerText = `${message} (数量 ${qty})`;
    toast.style.display = 'block';
    clearTimeout(toastTimeout);
    toastTimeout = setTimeout(() = { toast.style.display = 'none'; }, 2000);
}

function triggerVibration(type) {
    if (!navigator.vibrate) return;
    if (type === 'success') navigator.vibrate(50);               短く1回（手ごたえ用）
    if (type === 'send') navigator.vibrate([100, 100, 100]);     送信完了のブルッブルッ
    if (type === 'error') navigator.vibrate([500]);              エラーの長く1回
}

async function requestWakeLock() {
    try {
        if ('wakeLock' in navigator) {
            wakeLock = await navigator.wakeLock.request('screen');
        }
    } catch (err) {
        console.log(スリープ防止はサポートされていません);
    }
}

function releaseWakeLock() {
    if (wakeLock !== null) {
        wakeLock.release().then(() = { wakeLock = null; });
    }
}

function startScanner() {
    clearLocalStorage();
    
    const areaVal = document.getElementById('areaInput').value.trim();
    document.getElementById('currentAreaDisplay').innerText = areaVal  `エリア ${areaVal}`  エリア 未設定;

    const modal = document.getElementById('cameraModal');
    modal.style.display = 'flex'; 
    
    triggerVibration('success');

     画面の描画を待つため、0.1秒遅延させてカメラを確実に起動する
    setTimeout(() = {
        if (!html5QrCode) {
            html5QrCode = new Html5Qrcode(reader);
        }
        
        const config = { fps 10, qrbox { width 250, height 250 } };
        
        html5QrCode.start({ facingMode environment }, config, onScanSuccess)
            .then(() = { 
                isScanning = true; 
                requestWakeLock();
            })
            .catch((err) = {
                alert(カメラの起動に失敗しました。カメラのアクセス権限（許可）がオンになっているか確認してください。n詳細  + err);
                modal.style.display = 'none';
            });
    }, 100);
}

function stopScanner() {
    if (isScanning && html5QrCode) {
        html5QrCode.stop().then(() = {
            isScanning = false;
            document.getElementById('cameraModal').style.display = 'none';
            releaseWakeLock();
        }).catch(err = {
            document.getElementById('cameraModal').style.display = 'none';
            isScanning = false;
        });
    } else {
        document.getElementById('cameraModal').style.display = 'none';
    }
}

function onScanSuccess(decodedText, decodedResult) {
    const currentTime = Date.now();
    
    if (scannedDataMap.has(decodedText)) {
        const existingData = scannedDataMap.get(decodedText);
        if (currentTime - existingData.lastScanTime  500) return;  0.5秒未満は無視
        
        existingData.qty += 1;
        existingData.lastScanTime = currentTime;
        scannedDataMap.set(decodedText, existingData);
        
        saveToLocalStorage();
        triggerVibration('success');
        showToast(existingData.qrParts[0].trim(), existingData.qty);
        return;
    }

    try {
        const parts = decodedText.split(',');
        if (parts.length  3) return; 

        const orderNo = parts[0].trim();
        const areaVal = document.getElementById('areaInput').value.trim();

        scannedDataMap.set(decodedText, {
            qrParts parts,
            area areaVal,
            lastScanTime currentTime,
            qty 1
        });

        saveToLocalStorage();
        updateDataCount();
        triggerVibration('success');
        showToast(orderNo, 1);

    } catch (error) {
        console.error(解析エラー, error);
    }
}

async function sendDataToGAS() {
    if (scannedDataMap.size === 0) {
        alert(送信するデータがありません。);
        return;
    }

     即座の手ごたえフィードバック
    triggerVibration('success');

    const btnSend = document.getElementById('btnSend');
    const sysMsg = document.getElementById('systemMessage');
    
     視覚的な手ごたえ（ボタンの色と文字を変える）
    btnSend.disabled = true;
    btnSend.innerText = ⏳ 送信中...;
    btnSend.style.backgroundColor = #555; 
    sysMsg.innerText = ;

    const selectedMode = document.getElementById('modeSelect').value;
    const payload = [];
    
    scannedDataMap.forEach((value, key) = {
        const orderNo = value.qrParts[0].trim();
        const custName = value.qrParts[1].trim();
        const area = value.area;
        const qty = value.qty;

        for (let i = 2; i  value.qrParts.length; i++) {
            const rowNo = value.qrParts[i].trim();
            if (rowNo !== ) {
                payload.push({ order orderNo, name custName, row rowNo, area area, qty qty, mode selectedMode });
            }
        }
    });

    try {
        const response = await fetch(GAS_API_URL, {
            method 'POST',
            headers { 'Content-Type' 'textplain' },
            body JSON.stringify(payload),
            redirect 'follow'
        });

        const result = await response.json();

        if (result.status === success) {
            triggerVibration('send'); 
            alert(送信が完了しました！);
        } else {
            triggerVibration('error');
            sysMsg.innerText = エラー  + result.message;
        }
    } catch (error) {
        triggerVibration('error');
        sysMsg.innerText = 通信エラーが発生しました。電波の良い場所で再度お試しください。;
    } finally {
        btnSend.disabled = false;
        btnSend.innerText = スプレッドシートへ送信;
        btnSend.style.backgroundColor = #00c853;
    }
}