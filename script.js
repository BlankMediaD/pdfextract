document.getElementById('parse').addEventListener('click', () => {
    const data = document.getElementById('data').value;
    const regex = /^(\d+)\s(.*?)\n(.*?)(?:Ph (.*?))?(?:Email (.*?))?(?:Web (.*?))?(?:Rep\.(.*?))?(?:NB (.*?))(?=\n^\d+\s|$(?![\r\n]))/gsm;

    let match;
    const parsedData = [];
    let maxReps = 0;

    while ((match = regex.exec(data)) !== null) {
        const reps = match[7] ? match[7].trim().split('\n').map(r => r.trim()) : [];
        const parsedReps = reps.map(r => {
            const parts = r.split(' - ');
            const name = parts[0];
            const rest = parts.length > 1 ? parts[1] : '';
            const desMatch = rest.match(/(.*?)(\s\d+.*)?$/);
            const designation = desMatch ? desMatch[1].trim() : rest;
            const number = desMatch && desMatch[2] ? desMatch[2].trim() : '';
            return { name, designation, number };
        });

        if (parsedReps.length > maxReps) {
            maxReps = parsedReps.length;
        }

        parsedData.push({
            entryNum: match[1],
            companyName: match[2],
            address: match[3].replace(/\n/g, ' ').trim(),
            phone: match[4] ? match[4].replace(/\n/g, '').trim() : '',
            email: match[5] ? match[5].replace(/\n/g, '').trim() : '',
            website: match[6] ? match[6].replace(/\n/g, '').trim() : '',
            representatives: parsedReps,
            description: match[8] ? match[8].replace(/\n/g, ' ').trim() : ''
        });
    }

    const output = document.getElementById('output').getElementsByTagName('tbody')[0];
    const header = document.getElementById('output').getElementsByTagName('thead')[0].getElementsByTagName('tr')[0];
    output.innerHTML = '';
    header.innerHTML = '<th>Entry No.</th><th>Company Name</th><th>Address</th><th>Phone</th><th>Email</th><th>Website</th><th>Description</th>';

    for (let i = 1; i <= maxReps; i++) {
        header.innerHTML += `<th>Rep ${i} Name</th><th>Rep ${i} Designation</th><th>Rep ${i} Number</th>`;
    }

    parsedData.forEach(data => {
        const row = output.insertRow();
        row.insertCell().textContent = data.entryNum;
        row.insertCell().textContent = data.companyName;
        row.insertCell().textContent = data.address;
        row.insertCell().textContent = data.phone;
        row.insertCell().textContent = data.email;
        row.insertCell().textContent = data.website;
        row.insertCell().textContent = data.description;
        for (let i = 0; i < maxReps; i++) {
            const rep = data.representatives[i];
            row.insertCell().textContent = rep ? rep.name : '';
            row.insertCell().textContent = rep ? rep.designation : '';
            row.insertCell().textContent = rep ? rep.number : '';
        }
    });
});

document.getElementById('export').addEventListener('click', () => {
    const table = document.getElementById('output');
    const wb = XLSX.utils.table_to_book(table, {sheet:"Sheet JS"});
    XLSX.writeFile(wb, 'output.xlsx');
});
