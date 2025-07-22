document.getElementById('parse').addEventListener('click', () => {
    const data = document.getElementById('data').value;
    const regex = /^\d+\s/m;
    const entryNumbers = data.match(new RegExp(regex.source, 'g')).map(s => s.trim());
    const entries = data.split(regex).filter(Boolean);
    const output = document.getElementById('output').getElementsByTagName('tbody')[0];
    const header = document.getElementById('output').getElementsByTagName('thead')[0].getElementsByTagName('tr')[0];
    output.innerHTML = '';
    header.innerHTML = '<th>Entry No.</th><th>Company Name</th><th>Address</th><th>Phone</th><th>Email</th><th>Website</th><th>Description</th>';

    let maxReps = 0;

    const parsedData = entries.map((entry, i) => {
        const lines = entry.trim().split('\n');
        let companyName = lines[0].trim();
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

        let repStarted = false;
        while (j < lines.length) {
            const line = lines[j].trim();
            if (line.startsWith('Ph ')) {
                phone += line.substring(3).trim() + ' ';
            } else if (line.startsWith('Email ')) {
                email += line.substring(6).trim() + ' ';
            } else if (line.startsWith('Web ')) {
                website += line.substring(4).trim() + ' ';
            } else if (line.startsWith('Rep.')) {
                repStarted = true;
                const repLine = line.substring(5).trim();
                representatives.push(repLine);
            } else if (line.startsWith('NB ')) {
                repStarted = false;
                description += line.substring(3).trim() + ' ';
            } else if (repStarted) {
                representatives.push(line);
            } else {
                description += line + ' ';
            }
            j++;
        }

        const parsedReps = representatives.map(r => {
            const parts = r.split(' - ');
            const name = parts[0];
            const rest = parts.length > 1 ? parts[1] : '';
            const match = rest.match(/(.*?)(\s\d+.*)?$/);
            const designation = match ? match[1].trim() : rest;
            const number = match && match[2] ? match[2].trim() : '';
            return { name, designation, number };
        });

        if (parsedReps.length > maxReps) {
            maxReps = parsedReps.length;
        }

        return {
            entryNum: entryNumbers[i],
            companyName,
            address: address.trim(),
            phone: phone.trim(),
            email: email.trim(),
            website: website.trim(),
            representatives: parsedReps,
            description: description.trim(),
        };
    });

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
        data.representatives.forEach(rep => {
            row.insertCell().textContent = rep.name;
            row.insertCell().textContent = rep.designation;
            row.insertCell().textContent = rep.number;
        });
    });
});

document.getElementById('export').addEventListener('click', () => {
    const table = document.getElementById('output');
    const wb = XLSX.utils.table_to_book(table, {sheet:"Sheet JS"});
    XLSX.writeFile(wb, 'output.xlsx');
});
