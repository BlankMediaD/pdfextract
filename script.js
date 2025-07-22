document.getElementById('parse').addEventListener('click', () => {
    const data = document.getElementById('data').value;
    const regex = /^\d+\s/m;
    const entryNumbers = data.match(new RegExp(regex.source, 'g')).map(s => s.trim());
    const entries = data.split(regex).filter(Boolean);
    const output = document.getElementById('output').getElementsByTagName('tbody')[0];
    output.innerHTML = '';

    for (let i = 0; i < entries.length; i++) {
        const entry = entries[i];
        const entryNum = entryNumbers[i];
        const lines = entry.trim().split('\n');
        const companyName = lines[0].trim();
        let address = '';
        let phone = '';
        let email = '';
        let website = '';
        let representatives = [];
        let description = '';

        let j = 1;
        while (j < lines.length && !lines[j].startsWith('Ph ') && !lines[j].startsWith('Email ') && !lines[j].startsWith('Web ') && !lines[j].startsWith('Rep.')) {
            address += lines[j].trim() + ' ';
            j++;
        }

        while (j < lines.length) {
            if (lines[j].startsWith('Ph ')) {
                phone = lines[j].substring(3).trim();
            } else if (lines[j].startsWith('Email ')) {
                email = lines[j].substring(6).trim();
            } else if (lines[j].startsWith('Web ')) {
                website = lines[j].substring(4).trim();
            } else if (lines[j].startsWith('Rep.')) {
                let repLines = [];
                repLines.push(lines[j].substring(5).trim());
                j++;
                while(j < lines.length && !lines[j].startsWith('NB ')) {
                    repLines.push(lines[j].trim());
                    j++;
                }
                representatives = repLines.map(l => {
                    const parts = l.split(' - ');
                    const name = parts[0];
                    const rest = parts.length > 1 ? parts[1] : '';
                    const match = rest.match(/(.*?)(\s\d+)?$/);
                    const designation = match ? match[1] : rest;
                    const number = match && match[2] ? match[2].trim() : '';
                    return { name, designation, number };
                });
                j--;
            } else if (lines[j].startsWith('NB ')) {
                description = lines[j].substring(3).trim();
            }
            j++;
        }

        const row = output.insertRow();
        row.insertCell().textContent = entryNum;
        row.insertCell().textContent = companyName;
        row.insertCell().textContent = address.trim();
        row.insertCell().textContent = phone;
        row.insertCell().textContent = email;
        row.insertCell().textContent = website;
        row.insertCell().textContent = representatives.map(r => `${r.name} (${r.designation}) - ${r.number}`).join('; ');
        row.insertCell().textContent = description;
    }
});

document.getElementById('export').addEventListener('click', () => {
    const table = document.getElementById('output');
    const wb = XLSX.utils.table_to_book(table, {sheet:"Sheet JS"});
    XLSX.writeFile(wb, 'output.xlsx');
});
