document.getElementById('numbersForm').addEventListener('submit', function (event) {
    event.preventDefault();
    
    const numbersInput = document.getElementById('numbersInput').value;
    const quotes = document.getElementById('quotes').value;
    
    let numbersArray = numbersInput.split(' ').map(num => num.trim());
    if (quotes === 'yes') {
        numbersArray = numbersArray.map(num => `'${num}'`);
    }
    
    const result = numbersArray.join(', ');
    document.getElementById('resultNumbers').innerText = `Números: ${result}`;
});

document.getElementById('countForm').addEventListener('submit', function (event) {
    event.preventDefault();
    
    const countInput = document.getElementById('countInput').value;
    const numbersArray = countInput.split(/[\s,]+/).filter(Boolean); // separar por espacios o comas y filtrar vacíos
    
    const count = numbersArray.length;
    document.getElementById('resultCount').innerText = `Total de números: ${count}`;
});

document.getElementById('compareForm').addEventListener('submit', function (event) {
    event.preventDefault();
    
    const list1Input = document.getElementById('list1Input').value;
    const list2Input = document.getElementById('list2Input').value;
    
    const list1 = list1Input.split(/[\s,]+/).filter(Boolean);
    const list2 = list2Input.split(/[\s,]+/).filter(Boolean);
    
    const lengthComparison = list1.length > list2.length ? 'Lista 1 es más larga' : 'Lista 2 es más larga';
    
    const missingInList1 = list2.filter(num => !list1.includes(num));
    const missingInList2 = list1.filter(num => !list2.includes(num));
    
    const differencesCount = missingInList1.length + missingInList2.length;
    
    const result = `${lengthComparison}\nFaltan en Lista 1: ${missingInList1.join(', ') || 'Ninguno'}\nFaltan en Lista 2: ${missingInList2.join(', ') || 'Ninguno'}\nTotal de diferencias: ${differencesCount}`;
    document.getElementById('resultCompare').innerText = result;
});
