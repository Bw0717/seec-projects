let personnel = [];
let productionLines = [];
let stations = [];
let currentPersonId = null;

document.getElementById('personnelForm').addEventListener('submit', function(e) {
    e.preventDefault();
    const name = document.getElementById('name').value;
    if (!name.trim()) return;
    
    const person = {
        id: Date.now().toString(),
        name: name,
        stationSettings: {},
        onLeave: false
    };
    
    personnel.push(person);
    updatePersonnelList();
    this.reset();
});

document.getElementById('lineForm').addEventListener('submit', function(e) {
    e.preventDefault();
    const lineName = document.getElementById('lineName').value;
    if (!lineName.trim()) return;
    
    const line = {
        id: Date.now().toString(),
        name: lineName,
        stations: []
    };
    
    productionLines.push(line);
    updateLineList();
    updateLineSelect();
    this.reset();
});

document.getElementById('stationForm').addEventListener('submit', function(e) {
    e.preventDefault();
    
    const lineId = document.getElementById('lineSelect').value;
    const stationName = document.getElementById('stationName').value;
    
    if (!lineId || !stationName.trim()) {
        alert('請填寫完整工作站資訊');
        return;
    }
    
    const station = {
        id: Date.now().toString(),
        name: stationName,
        highPriority: document.getElementById('highPriority').checked,
        priorityAssign: document.getElementById('priorityAssign').checked,
        lineId: lineId
    };
    
    stations.push(station);
    
    const line = productionLines.find(l => l.id === lineId);
    if (line) {
        line.stations.push(station.id);
        alert(`工作站 ${stationName} 新增成功！`);
    }
    
    updateLineList();
    
    document.getElementById('stationName').value = '';
    document.getElementById('highPriority').checked = false;
    document.getElementById('priorityAssign').checked = false;
});

document.getElementById('stationForm').innerHTML = `
    <div class="form-group">
        <label for="lineSelect">選擇產線：</label>
        <select id="lineSelect" required>
            <option value="">請選擇產線</option>
        </select>
    </div>
    <div class="form-group">
        <label for="stationName">工作站名稱：</label>
        <input type="text" id="stationName" required>
    </div>
    <div class="form-group">
        <label class="checkbox-label">
            <input type="checkbox" id="highPriority">
            <span>不受人數限制</span>
        </label>
    </div>
    <div class="form-group">
        <label class="checkbox-label">
            <input type="checkbox" id="priorityAssign">
            <span>人員不足時優先分配</span>
        </label>
    </div>
    <button type="submit" class="btn-primary">新增工作站</button>
`;

function updateLineSelect() {
    const lineSelect = document.getElementById('lineSelect');
    lineSelect.innerHTML = '<option value="">請選擇產線</option>';
    
    productionLines.forEach(line => {
        const option = document.createElement('option');
        option.value = line.id;
        option.textContent = line.name;
        lineSelect.appendChild(option);
    });
}

function updateLineList() {
    const lineList = document.getElementById('lineList');
    lineList.innerHTML = '';
    
    productionLines.forEach(line => {
        const div = document.createElement('div');
        div.className = 'list-item';
        div.innerHTML = `
            <h3>${line.name}</h3>
            <div class="line-stations">
                <h4>工作站：</h4>
                <div class="stations-container">
                    ${line.stations.map(stationId => {
                        const station = stations.find(s => s.id === stationId);
                        return station ? 
                            `<div class="station-item">
                                <span class="station-name">${station.name}</span>
                                <span class="priority-badge">
                                    ${station.highPriority ? '🔒' : ''} 
                                    ${station.priorityAssign ? '⭐' : ''}
                                </span>
                                <div class="station-actions">
                                    <span class="edit-icon" onclick="editStation('${station.id}')">✏️</span>
                                    <span class="delete-icon" onclick="deleteStation('${station.id}')">🗑️</span>
                                </div>
                            </div>` : '';
                    }).join('')}
                </div>
            </div>
            <button onclick="deleteLine('${line.id}')" class="btn-danger">刪除產線</button>
        `;
        lineList.appendChild(div);
    });
    autoSaveSettings();
}

function deleteStation(id) {
    if (confirm('確定要刪除此工作站嗎？')) {
        const station = stations.find(s => s.id === id);
        if (station) {
            const line = productionLines.find(l => l.id === station.lineId);
            if (line) {
                line.stations = line.stations.filter(stationId => stationId !== id);
            }
            
            stations = stations.filter(station => station.id !== id);
            
            updateLineList();
            autoSaveSettings();
        }
    }
}

function deleteLine(id) {
    if (confirm('確定要刪除此產線嗎？這將同時刪除所有相關工作站。')) {
        stations = stations.filter(station => station.lineId !== id);
        productionLines = productionLines.filter(line => line.id !== id);
        updateLineList();
        updateStationOptions();
        updateLineSelect();
    }
}

function updateStationList() {
    const stationList = document.getElementById('stationList');
    stationList.innerHTML = '';
    
    stations.forEach(station => {
        const line = productionLines.find(l => l.id === station.lineId);
        const div = document.createElement('div');
        div.className = 'list-item';
        div.innerHTML = `
            <h3>${station.name}</h3>
            <p>所屬產線: ${line ? line.name : '未知'}</p>
            <button onclick="deleteStation(${station.id})">刪除</button>
        `;
        stationList.appendChild(div);
    });
}

function updatePersonnelList() {
    const personnelList = document.getElementById('personnelList');
    personnelList.innerHTML = '';
    
    personnel.forEach(person => {
        const div = document.createElement('div');
        div.className = 'list-item';
        div.innerHTML = `
            <h3>
                ${person.name}
                ${person.onLeave ? '<span class="leave-badge">休假中</span>' : ''}
            </h3>
            <div class="person-actions">
                <label class="checkbox-label">
                    <input type="checkbox" 
                           ${person.onLeave ? 'checked' : ''} 
                           onchange="toggleLeave('${person.id}')">
                    <span>休假</span>
                </label>
                <button onclick="openSettings('${person.id}')" class="btn-primary">工作站設定</button>
                <button onclick="deletePerson('${person.id}')" class="btn-danger">刪除</button>
            </div>
        `;
        personnelList.appendChild(div);
    });
    autoSaveSettings();
}

function getPersonStationSummary(person) {
    const enabledStations = [];
    const preferences = [];
    
    for (const [stationId, setting] of Object.entries(person.stationSettings)) {
        const station = stations.find(s => s.id === stationId);
        if (station && setting.ability) {
            enabledStations.push(station.name);
            if (setting.preference !== 'normal') {
                preferences.push(`${station.name}(${setting.preference})`);
            }
        }
    }
    
    let summary = `<p>可工作站點: ${enabledStations.length ? enabledStations.join(', ') : '無'}</p>`;
    if (preferences.length) {
        summary += `<p>特殊偏好: ${preferences.join(', ')}</p>`;
    }
    return summary;
}

function deletePerson(id) {
    const person = personnel.find(p => p.id === id);
    if (person && confirm(`確定要刪除 ${person.name} 嗎？`)) {
        personnel = personnel.filter(person => person.id !== id);
        updatePersonnelList();
    }
}

function openSettings(personId) {
    currentPersonId = personId;
    const modal = document.getElementById('settingsModal');
    const settingsDiv = document.getElementById('stationSettings');
    const person = personnel.find(p => p.id === personId);
    
    settingsDiv.innerHTML = '';
    
    productionLines.forEach(line => {
        const lineStations = stations.filter(station => station.lineId === line.id);
        if (lineStations.length > 0) {
            const lineDiv = document.createElement('div');
            lineDiv.className = 'line-settings-group';
            lineDiv.innerHTML = `
                <h3 class="line-name">${line.name}</h3>
                <div class="line-stations-settings">
                    ${lineStations.map(station => {
                        const setting = person.stationSettings[station.id] || { ability: false, preference: 'normal' };
                        return `
                            <div class="station-setting">
                                <h4>${station.name}</h4>
                                <div class="setting-controls">
                                    <div class="ability-toggle">
                                        <label class="switch">
                                            <input type="checkbox" class="station-ability" 
                                                   data-station="${station.id}"
                                                   ${setting.ability ? 'checked' : ''}>
                                            <span class="slider"></span>
                                        </label>
                                        <span>具備工作能力</span>
                                    </div>
                                    <div class="preference-select">
                                        <label title="工作偏好程度">
                                            偏好：
                                            <select class="station-preference" data-station="${station.id}">
                                                <option value="high" ${setting.preference === 'high' ? 'selected' : ''}>高</option>
                                                <option value="normal" ${setting.preference === 'normal' ? 'selected' : ''}>一般</option>
                                                <option value="low" ${setting.preference === 'low' ? 'selected' : ''}>低</option>
                                            </select>
                                        </label>
                                    </div>
                                </div>
                            </div>
                        `;
                    }).join('')}
                </div>
            `;
            settingsDiv.appendChild(lineDiv);
        }
    });
    
    modal.style.display = 'block';
}

function saveSettings() {
    const person = personnel.find(p => p.id === currentPersonId);
    if (!person) return;
    
    person.stationSettings = {};
    
    document.querySelectorAll('.station-setting').forEach(div => {
        const stationId = div.querySelector('.station-ability').dataset.station;
        person.stationSettings[stationId] = {
            ability: div.querySelector('.station-ability').checked,
            preference: div.querySelector('.station-preference').value
        };
    });
    
    document.getElementById('settingsModal').style.display = 'none';
    updatePersonnelList();
    autoSaveSettings();
    
    alert(`${person.name} 的工作站設定已儲存成功！`);
}

document.querySelector('.close').onclick = function() {
    document.getElementById('settingsModal').style.display = 'none';
}

window.onclick = function(event) {
    const modal = document.getElementById('settingsModal');
    if (event.target === modal) {
        modal.style.display = 'none';
    }
}

function saveToJSON() {
    const settings = {
        personnel: personnel,
        productionLines: productionLines,
        stations: stations,
        version: '1.0',
        savedAt: new Date().toISOString()
    };

    const blob = new Blob([JSON.stringify(settings, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `workplace_settings_${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function loadFromJSON(input) {
    const file = input.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const settings = JSON.parse(e.target.result);
                
                if (!settings.personnel || !settings.productionLines || !settings.stations) {
                    throw new Error('無效的設定檔格式');
                }

                personnel = settings.personnel;
                productionLines = settings.productionLines;
                stations = settings.stations;

                updateLineList();
                updateLineSelect();
                updatePersonnelList();

                alert('設定已成功載入！');
            } catch (error) {
                alert('載入設定時發生錯誤：' + error.message);
            }
        };
        reader.readAsText(file);
    }
    input.value = '';
}

document.addEventListener('DOMContentLoaded', function() {
    const savedSettings = localStorage.getItem('workplaceSettings');
    if (savedSettings) {
        try {
            const settings = JSON.parse(savedSettings);
            personnel = settings.personnel;
            productionLines = settings.productionLines;
            stations = settings.stations;
            updateLineList();
            updateLineSelect();
            updatePersonnelList();
        } catch (error) {
            console.error('載入本地設定時發生錯誤：', error);
        }
    }
});

function autoSaveSettings() {
    const settings = {
        personnel: personnel,
        productionLines: productionLines,
        stations: stations,
        savedAt: new Date().toISOString()
    };
    localStorage.setItem('workplaceSettings', JSON.stringify(settings));
}

function editStation(stationId) {
    const station = stations.find(s => s.id === stationId);
    if (!station) return;

    const modal = document.createElement('div');
    modal.id = 'editStationModal';
    modal.className = 'modal';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h2>編輯工作站</h2>
                <span class="close" onclick="closeEditModal()">&times;</span>
            </div>
            <div class="modal-body">
                <div class="form-group">
                    <label for="editStationName">工作站名稱：</label>
                    <input type="text" id="editStationName" value="${station.name}" required>
                </div>
                <div class="form-group">
                    <label for="editStationOrder">順序：</label>
                    <input type="number" id="editStationOrder" value="${station.order || 0}" min="0" required>
                </div>
                <div class="form-group">
                    <label class="checkbox-label">
                        <input type="checkbox" id="editHighPriority" ${station.highPriority ? 'checked' : ''}>
                        <span>不受人數限制</span>
                    </label>
                </div>
                <div class="form-group">
                    <label class="checkbox-label">
                        <input type="checkbox" id="editPriorityAssign" ${station.priorityAssign ? 'checked' : ''}>
                        <span>人員不足時優先分配</span>
                    </label>
                </div>
            </div>
            <div class="modal-footer">
                <button onclick="saveStationEdit('${stationId}')" class="btn-primary">儲存</button>
            </div>
        </div>
    `;
    document.body.appendChild(modal);

    modal.onclick = function(event) {
        if (event.target === modal) {
            closeEditModal();
        }
    }

    modal.style.display = 'block';
}

function closeEditModal() {
    const modal = document.getElementById('editStationModal');
    if (modal) {
        document.body.removeChild(modal);
    }
}

function saveStationEdit(stationId) {
    const station = stations.find(s => s.id === stationId);
    if (!station) return;

    const newName = document.getElementById('editStationName').value;
    const newOrder = parseInt(document.getElementById('editStationOrder').value);
    const newHighPriority = document.getElementById('editHighPriority').checked;
    const newPriorityAssign = document.getElementById('editPriorityAssign').checked;

    if (!newName.trim()) {
        alert('請輸入工作站名稱');
        return;
    }

    station.name = newName;
    station.order = newOrder;
    station.highPriority = newHighPriority;
    station.priorityAssign = newPriorityAssign;

    updateLineList();
    autoSaveSettings();
    closeEditModal();
    alert('工作站更新成功！');
}

function executeTask() {
    if (confirm('確定要執行排程嗎？')) {
        const data = {
            personnel: personnel,
            productionLines: productionLines,
            stations: stations,
            timestamp: new Date().toISOString()
        };

        document.body.classList.add('loading');
        
        fetch('https://seec-auto.vercel.app/api/schedule', {  // 確保這是您的 Flask 服務器地址
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(data)
        })
        .then(response => {
            if (!response.ok) {
                throw new Error('排程計算失敗');
            }
            return response.json();
        })
        .then(result => {
            localStorage.setItem('scheduleResults', JSON.stringify(result));
            window.location.href = 'result.html';
        })
        .catch(error => {
            alert('執行排程時發生錯誤：' + error.message);
        })
        .finally(() => {
            document.body.classList.remove('loading');
        });
    }
}

function toggleLeave(personId) {
    const person = personnel.find(p => p.id === personId);
    if (person) {
        person.onLeave = !person.onLeave;
        updatePersonnelList();
    }
}

updateLineSelect(); 