document.getElementById('settings-form').addEventListener('submit', function (e) {
    e.preventDefault();
    const startTime = document.getElementById('startTime').value;
    const duration = parseInt(document.getElementById('duration').value);
    const breakTime = parseInt(document.getElementById('breakTime').value);
    const excelFile = document.getElementById('csvFile').files[0];

    const reader = new FileReader();
    reader.onload = async function (e) {

        // tableタグを削除
        const scheduleDiv = document.getElementById('schedule');
        scheduleDiv.innerHTML = '';


        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const studios = parseExcelAndInitializeData(firstSheet);

        for (let s of studios) {
            console.log(s.name, s.score_vec);
        }
        // 一つでも areSupervisorsNonConflicting が true ならば、スケジュール調整が不可能
        {

            let isAllScoreVecOne = false;
            let impossibleStudios = [];

            for (let studio of studios) {
                if (studio.score_vec.every(score => score > 0)) {
                    isAllScoreVecOne = true;
                    impossibleStudios.push(studio.name);
                    break;
                }

            }
            if (isAllScoreVecOne) {
                alert('スケジュール調整が不可能です。');
                document.querySelector('#info').innerHTML = `すべての審査員が含まれるスタジオ研究室: ${impossibleStudios.join(', ')}`;
                return;
            }
        }


        // studiosから学生の発表数を取得
        const studentsCount = studios.map(s => s.students.length).reduce((acc, val) => acc + val, 0);
        document.querySelector('#info').innerHTML = `研究室数: ${studios.length}, 学生数: ${studentsCount}`;
        // const schedule = generateSchedule(studios, startTime, duration, breakTime);
        // if (schedule) {
        //     displaySchedule(schedule);
        // }
        for (let i = 0; i < 10000; i++) {
            const schedule = await generateSchedule(studios, startTime, duration, breakTime, -1);
            if (schedule.is_succeeded) {
                console.log(`[${i}]:スケジュール調整に成功しました。`);

                displaySchedule(schedule.rooms);
                document.querySelector('#exportExcel').disabled = false;
                break;
            }
            else {
                document.querySelector('#exportExcel').disabled = true;
                document.querySelector('#schedule').innerHTML = 'スケジュール調整に失敗しました。もう一度「スケジュールを作成」ボタンを押してください。';
                console.log(`[${i}]:スケジュール調整に失敗しました。`);
                document.querySelector('#timetable_info').innerHTML = `スケジュール調整回数: ${i + 1}`;
            }
        }
    };
    reader.readAsArrayBuffer(excelFile);
});

function parseExcelAndInitializeData(sheet) {
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const headers = json[0];
    const studios = {};
    const supervisors = new Set();


    // headersの要素名をチェック
    if (headers.indexOf('no') === -1 ||
        headers.indexOf('name') === -1 ||
        headers.indexOf('name_en') === -1 ||
        headers.indexOf('studio') === -1 ||
        headers.indexOf('title') === -1 ||
        headers.indexOf('supervisor') === -1 ||
        headers.indexOf('sub1') === -1 ||
        headers.indexOf('sub2') === -1
    ) {
        alert('必要なヘッダーが見つかりません。サンプルファイルのヘッダー名と読み込みファイルのヘッダー名を確認してください。');
        return
    }

    let profs = [];
    for (let i = 1; i < json.length; i++) {
        const currentline = json[i];
        if (currentline.length === headers.length) {
            const student = {};
            for (let j = 0; j < headers.length; j++) {
                student[headers[j]] = currentline[j];
            }
            supervisors.add(student['supervisor']);
            supervisors.add(student['sub1']);
            supervisors.add(student['sub2']);
            profs.push(student['supervisor']);
            profs.push(student['sub1']);
            profs.push(student['sub2']);

            if (!studios[student['studio']]) {
                studios[student['studio']] = {
                    name: student['studio'],
                    score_vec: [],
                    students: []
                };
            }
            studios[student['studio']].students.push({
                name: student['name'],
                name_en: student['name_en'],
                id: student['no'],
                studioName: student['studio'],
                title: student['title'],
                supervisor: student['supervisor'],
                subsupervisors: [student['sub1'], student['sub2']],
                time: { start: '', end: '' },
                score_vec: []
            });
        }
    }

    // professorの重複を削除
    profs = Array.from(new Set(profs));
    // profsをalphabetical orderに並び替え
    profs.sort();


    const supervisorArray = Array.from(supervisors);
    for (let studioName in studios) {
        const studio = studios[studioName];
        studio.score_vec = new Array(supervisorArray.length).fill(0);

        for (let student of studio.students) {
            student.score_vec = new Array(supervisorArray.length).fill(0);
            const supervisorIndex = supervisorArray.indexOf(student.supervisor);
            const subsupervisor1Index = supervisorArray.indexOf(student.subsupervisors[0]);
            const subsupervisor2Index = supervisorArray.indexOf(student.subsupervisors[1]);

            if (supervisorIndex !== -1) studio.score_vec[supervisorIndex]++;
            if (subsupervisor1Index !== -1) studio.score_vec[subsupervisor1Index]++;
            if (subsupervisor2Index !== -1) studio.score_vec[subsupervisor2Index]++;

            if (supervisorIndex !== -1) student.score_vec[supervisorIndex]++;
            if (subsupervisor1Index !== -1) student.score_vec[subsupervisor1Index]++;
            if (subsupervisor2Index !== -1) student.score_vec[subsupervisor2Index]++;
        }
    }

    return Object.values(studios);
}





function areSupervisorsNonConflicting(studio1, studio2) {
    const supervisors1 = new Set(studio1.students.map(student => student.supervisor));
    const supervisors2 = new Set(studio2.students.map(student => student.supervisor));
    const subsupervisors1 = new Set(studio1.students.flatMap(student => student.subsupervisors));
    const subsupervisors2 = new Set(studio2.students.flatMap(student => student.subsupervisors));

    for (let supervisor of supervisors1) {
        if (subsupervisors2.has(supervisor)) {
            return false; // 研究室1の主査が研究室2の副査に含まれている
        }
    }
    for (let supervisor of supervisors2) {
        if (subsupervisors1.has(supervisor)) {
            return false; // 研究室2の主査が研究室1の副査に含まれている
        }
    }
    return true;
}





function getRandomStudio(otherStudios) {
    if (otherStudios.length < 1) {
        return null;
    } else {
        // ランダムにインデックスを選び、要素を削除しつつ取得
        let randomIndex = Math.floor(Math.random() * otherStudios.length);
        let ret_studio = otherStudios.splice(randomIndex, 1)[0]; // 配列から削除して取得
        return ret_studio;
    }
}

function getRandomStudent(students) {
    if (students.length < 1) {
        return null;
    } else {
        // ランダムにインデックスを選び、要素を削除しつつ取得
        let randomIndex = Math.floor(Math.random() * students.length);
        let ret_student = students.splice(randomIndex, 1)[0]; // 配列から削除して取得
        return ret_student;
    }
}


function findRandomNonConflictingStudio(studio, studios) {

    // まず、studio と主査および副査が互いに重ならないスタジオをすべて選出
    let nonConflictingStudios = studios.filter(s =>
        areSupervisorsNonConflicting(studio, s)
    );

    // ランダムに1つのスタジオを選択して返す
    if (nonConflictingStudios.length > 0) {
        let randomIndex = Math.floor(Math.random() * nonConflictingStudios.length);
        return nonConflictingStudios[randomIndex];
    } else {
        // 該当するスタジオがない場合は null を返す
        return null;
    }
}






async function generateSchedule(studios, startTime, duration, breakTime, conflict_threshold = -1) {
    let room1 = [];
    let room2 = [];
    let otherStudios = studios.slice();
    let allStudios = studios.slice();

    let time1 = startTime;
    let time2 = startTime;

    // 最初に一つ分の研究室を取り出す。
    let studio1 = getRandomStudio(otherStudios);
    let students1 = studio1.students.slice();
    let studio2 = getRandomStudio(otherStudios);
    let students2 = studio2.students.slice();

    // console.log("スケジュール調整スタート", studio1.name, studio2.name, otherStudios.length);

    let tempRoom1 = [];
    let tempRoom2 = [];


    let tempTime1 = time1;
    let tempTime2 = time2;

    let conflict_count = 0;

    let was_lunch_assigned = false;
    let break_time_start = document.querySelector('#breakTime_start').value;
    // メインのアサイン処理ループ
    while (otherStudios.length >= 1) { // otherStudiosが1つ以上である限りはループ

        // console.log("   ", otherStudios.length, students1.length, students2.length);
        let student1 = null;
        let student2 = null;

        // studio1の学生がいない場合
        if (students1.length === 0) {
            // studio1に別の研究室を割り当てる。ただし studio2は除き、審査も重複しない研究室とする
            studio1 = getRandomStudio(otherStudios);
            students1 = studio1.students.slice();
        }

        // 1つ目の研究室の学生をランダムに取り出す。取り出された要素は削除される。
        student1 = getRandomStudent(students1);

        // もし students2にもう学生がいない場合、別の研究室に変更する
        if (students2.length === 0) {
            if (otherStudios.length === 0) {
                // student1を students1に戻す
                students1.push(student1);
                break;
            }
            studio2 = getRandomStudio(otherStudios);
            students2 = studio2.students.slice();
        }

        let flg_conflict = false;

        // 2つ目の研究室の学生を取り出す。ただし、審査員重複がない学生を取り出す

        student2 = students2.find(s => !isDuplicated(student1.score_vec, s.score_vec));

        if (!student2) {
            return { is_succeeded: false };
        }
        else if (student2) {
            // console.log("       ", student1.score_vec, student2.score_vec);
            students2 = students2.filter(s => s !== student2);

            tempRoom1.push({
                ...student1,
                time: { start: tempTime1, end: addMinutes(tempTime1, duration) },
                room: 1, conflict: flg_conflict
            });
            tempRoom2.push({
                ...student2,
                time: { start: tempTime2, end: addMinutes(tempTime2, duration) },
                room: 2, conflict: flg_conflict
            });
            tempTime1 = addMinutes(tempTime1, duration);
            tempTime2 = addMinutes(tempTime2, duration);
        }

        // Lunch Break処理：もし students1 が空になったいて、かつ tempTime1が12:00を過ぎていれば
        if (students1.length === 0 && tempTime1 >= break_time_start &&
            !was_lunch_assigned
        ) {
            // students2が残っていればすべてroom2に割り当てる
            for (let s2 of students2) {
                console.log(s2.name);
                tempRoom2.push({
                    ...s2,
                    time: { start: tempTime2, end: addMinutes(tempTime2, duration) },
                    room: 2, conflict: false
                });
                tempTime2 = addMinutes(tempTime2, duration);
            }
            students2 = [];
            tempRoom1.push({ lunchBreak: true, time: { start: tempTime1, end: addMinutes(tempTime2, breakTime) } });
            tempRoom2.push({ lunchBreak: true, time: { start: tempTime2, end: addMinutes(tempTime2, breakTime) } });
            tempTime2 = addMinutes(tempTime2, breakTime);
            tempTime1 = tempTime2;
            was_lunch_assigned = true;
        }
        else if (students2.length === 0 && tempTime2 >= break_time_start &&
            !was_lunch_assigned) {
            // students1が残っていればすべてroom1に割り当てる
            for (let s1 of students1) {
                tempRoom1.push({
                    ...s1,
                    time: { start: tempTime1, end: addMinutes(tempTime1, duration) },
                    room: 1, conflict: false
                });
                tempTime1 = addMinutes(tempTime1, duration);
            }
            students1 = [];
            tempRoom1.push({ lunchBreak: true, time: { start: tempTime1, end: addMinutes(tempTime1, breakTime) } });
            tempRoom2.push({ lunchBreak: true, time: { start: tempTime2, end: addMinutes(tempTime1, breakTime) } });
            tempTime1 = addMinutes(tempTime1, breakTime);
            tempTime2 = tempTime1;

            was_lunch_assigned = true;
        }
    }

    // console.log("残りの研究室", otherStudios.length);
    console.log("残りの学生数", students1.length, students2.length);
    console.log(students1, students2);
    if (students1.length > 0 && students2.length > 0) {
        return { is_succeeded: false };
    }
    else { // 普通に残りを
        for (let s1 of students1) {
            tempRoom1.push({
                ...s1,
                time: { start: tempTime1, end: addMinutes(tempTime1, duration) },
                room: 1, conflict: false
            });
            tempTime1 = addMinutes(tempTime1, duration);
        }
        for (let s2 of students2) {
            tempRoom2.push({
                ...s2,
                time: { start: tempTime2, end: addMinutes(tempTime2, duration) },
                room: 2, conflict: false
            });
            tempTime2 = addMinutes(tempTime2, duration);
        }
    }


    // otherStudiosにはもう一つだけスタジオが残っているのでその分を部屋1に追加する。ただし割当人数が多い部屋に追加する
    if (otherStudios.length > 0) {
        let room_assign = tempRoom1.length > tempRoom2.length ? 1 : 2;
        // console.log(`残りの研究室はroom[${room_assign}]に割り当てました`);
        for (let os of otherStudios) {
            for (let student of os.students) {
                if (room_assign === 1) {
                    tempRoom1.push({
                        ...student,
                        time: { start: tempTime1, end: addMinutes(tempTime1, duration) },
                        room: room_assign, conflict: false
                    });
                    tempTime1 = addMinutes(tempTime1, duration);
                }
                else {
                    tempRoom2.push({ ...student, time: { start: tempTime2, end: addMinutes(tempTime2, duration) }, room: room_assign, conflict: false });
                    tempTime2 = addMinutes(tempTime2, duration);
                }
            }
        }

    }


    room1 = room1.concat(tempRoom1);
    room2 = room2.concat(tempRoom2);
    time1 = tempTime1;
    time2 = tempTime2;


    // console.log(room1, room2, time1, time2);


    return { is_succeeded: true, rooms: [room1, room2] }

}




function isDuplicated(vec1, vec2) {
    for (let i = 0; i < vec1.length; i++) {
        if (vec1[i] && vec2[i]) {
            return true; // 重複あり
        }
    }
    return false; // 重複なし
}

function addMinutes(time, minsToAdd) {
    const [hours, minutes] = time.split(':').map(Number);
    const totalMinutes = hours * 60 + minutes + minsToAdd;
    const newHours = Math.floor(totalMinutes / 60).toString().padStart(2, '0');
    const newMinutes = (totalMinutes % 60).toString().padStart(2, '0');
    return `${newHours}:${newMinutes}`;
}

function displaySchedule(schedule) {
    const scheduleDiv = document.getElementById('schedule');
    scheduleDiv.innerHTML = '';
    const row = document.createElement('div');
    row.classList.add('row');

    schedule.forEach((roomSchedule, roomIndex) => {
        let presentation_count = 0;
        const col = document.createElement('div');
        col.classList.add('col-md-12', 'col-lg-6');
        const roomDiv = document.createElement('div');
        roomDiv.innerHTML = `<h3>部屋 ${roomIndex + 1} </h3>`;
        const table = document.createElement('table');
        table.classList.add('table', 'table-striped', 'table-hover', 'table-sm', 'align-middle');
        table.style.fontSize = '0.7rem';
        const thead = document.createElement('thead');
        thead.innerHTML = `
            <tr>
                <th>学生番号</th>
                <th>氏名</th>
                <th hidden>氏名（英語）</th>
                <th>所属研究室名</th>
                <th hidden>発表タイトル</th>
                <th>主査</th>
                <th>副査1</th>
                <th>副査2</th>
                <th>開始時間</th>
                <th>終了時間</th>
            </tr>
        `;
        table.appendChild(thead);
        const tbody = document.createElement('tbody');
        roomSchedule.forEach(item => {
            const row = document.createElement('tr');
            if (item.lunchBreak) {
                row.innerHTML = `<td class="text-center" colspan="8">Lunch Break: ${item.time.start} - ${item.time.end}</td>`;
            } else {
                row.innerHTML = `
                    <td>${item.id || ''}</td>
                    <td>${item.name || ''}</td>
                    <td hidden>${item.name_en || ''}</td>
                    <td>${item.studioName || ''}</td>
                    <td hidden>${item.title || ''}</td>
                    <td>${item.supervisor || ''}</td>
                    <td>${item.subsupervisors[0] || ''}</td>
                    <td>${item.subsupervisors[1] || ''}</td>
                    <td>${item.time.start}</td>
                    <td>${item.time.end}</td>
                `;
                presentation_count++;
                if (item.conflict) {
                    row.classList.add('table-danger');  // 重複がある場合にクラスを追加
                }
            }
            tbody.appendChild(row);
        });
        roomDiv.innerHTML += `<p>発表数: ${presentation_count}</p>`;
        table.appendChild(tbody);
        roomDiv.appendChild(table);
        col.appendChild(roomDiv);
        row.appendChild(col);
    });

    scheduleDiv.appendChild(row);
}


document.getElementById('exportExcel').addEventListener('click', function () {
    const wb = XLSX.utils.book_new();
    const tables = document.querySelectorAll('#schedule table');
    tables.forEach((table, index) => {
        const ws = XLSX.utils.table_to_sheet(table);
        XLSX.utils.book_append_sheet(wb, ws, `Room${index + 1}`);
    });
    XLSX.writeFile(wb, 'schedule.xlsx');
});
