document.getElementById("excelFile").addEventListener("change", handleFile, false);
document.getElementById("generateScript").addEventListener("click", generateSQLScript);

let parsedRows = [];

function handleFile(event) {
    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        processData(jsonData);
    };
    reader.readAsArrayBuffer(event.target.files[0]);
}

function processData(data) {
    const table = document.createElement("table");
    const header = `
    <tr>
      <th>Pk_IdVenta</th>
      <th>Total Volumen</th>
      <th>Diferencia Volumen</th>
      <th>Total Dinero</th>
      <th>Diferencia Dinero</th>
      <th>CantidadTotal</th>
      <th>ValorTotal</th>
    </tr>`;
    table.innerHTML = header;

    parsedRows = data.map((row, i) => {
        const current = {
            Pk_IdVenta: row["Pk_IdVenta"],
            VolumenInicial: parseFloat(row["VolumenInicial"]),
            VolumenFinal: parseFloat(row["VolumenFinal"]),
            DineroInicial: parseFloat(row["DineroInicial"]),
            DineroFinal: parseFloat(row["DineroFinal"]),
            CantidadTotal: parseFloat(row["CantidadTotal"]),
            ValorTotal: parseFloat(row["ValorTotal"]),
        };

        current.TotalVolumen = current.VolumenFinal - current.VolumenInicial;
        current.TotalDinero = current.DineroFinal - current.DineroInicial;

        const next = data[i + 1];
        if (next) {
            current.DifVolumen = parseFloat(next["VolumenInicial"]) - current.VolumenFinal;
            current.DifDinero = parseFloat(next["DineroInicial"]) - current.DineroFinal;
        } else {
            current.DifVolumen = 0;
            current.DifDinero = 0;
        }

        return current;
    });

    parsedRows.forEach((r) => {
        const tr = document.createElement("tr");

        const diffVolErr = Math.abs(r.DifVolumen) > 0.09;
        const diffDinErr = Math.abs(r.DifDinero) > 0.09;
        const totalVolErr = Math.abs(r.TotalVolumen - r.CantidadTotal) > 0.0001;
        const totalDinErr = Math.abs(r.TotalDinero - r.ValorTotal) > 0.0001;

        tr.innerHTML = `
      <td>${r.Pk_IdVenta}</td>
      <td class="${totalVolErr ? "error" : ""}">${r.TotalVolumen.toFixed(2)}</td>
      <td class="${diffVolErr ? "error" : ""}">${r.DifVolumen.toFixed(2)}</td>
      <td class="${totalDinErr ? "error" : ""}">${r.TotalDinero.toFixed(2)}</td>
      <td class="${diffDinErr ? "error" : ""}">${r.DifDinero.toFixed(2)}</td>
      <td>${r.CantidadTotal.toFixed(2)}</td>
      <td>${r.ValorTotal.toFixed(2)}</td>
    `;
        table.appendChild(tr);
    });

    document.getElementById("tableContainer").innerHTML = "";
    table.classList.add("table", "table-bordered", "table-striped", "align-middle");
    document.getElementById("tableContainer").appendChild(table);

    const btn = document.getElementById("generateScript");
    btn.classList.remove("d-none");

}

function generateSQLScript() {
    let script = "";

    for (let i = 0; i < parsedRows.length; i++) {
        const current = parsedRows[i];
        const next = parsedRows[i + 1];

        let modificoSiguiente = false;

        // Diferencia Volumen
        if (next && Math.abs(current.DifVolumen) > 0.09) {
            const nuevoVolumenInicial = current.VolumenFinal;
            script += `
UPDATE [EstacionNSX].[dbo].[DetalleVentaCombustible]
SET VolumenInicial='${nuevoVolumenInicial.toFixed(2).replace(",", ".")}'
WHERE Fk_IdVenta='${next.Pk_IdVenta}';`;

            next.VolumenInicial = nuevoVolumenInicial;
            next.TotalVolumen = next.VolumenFinal - next.VolumenInicial;
            modificoSiguiente = true;
        }

        // Diferencia Dinero
        if (next && Math.abs(current.DifDinero) > 0.09) {
            const nuevoDineroInicial = Math.round(current.DineroFinal);
            script += `
UPDATE [EstacionNSX].[dbo].[DetalleVentaCombustible]
SET DineroInicial='${nuevoDineroInicial}'
WHERE Fk_IdVenta='${next.Pk_IdVenta}';`;

            next.DineroInicial = nuevoDineroInicial;
            next.TotalDinero = next.DineroFinal - next.DineroInicial;
            modificoSiguiente = true;
        }

        // Si se modificó la siguiente fila, evaluamos sus nuevos totales vs columnas originales
        if (next && modificoSiguiente) {
            // CantidadTotal
            if (Math.abs(next.TotalVolumen - next.CantidadTotal) > 0.0001) {
                script += `
UPDATE [EstacionNSX].[dbo].[Venta]
SET CantidadTotal='${next.TotalVolumen.toFixed(2).replace(",", ".")}'
WHERE Pk_IdVenta='${next.Pk_IdVenta}';`;
            }

            // ValorTotal
            if (Math.abs(next.TotalDinero - next.ValorTotal) > 0.0001) {
                script += `
UPDATE [EstacionNSX].[dbo].[Venta]
SET ValorTotal='${Math.round(next.TotalDinero)}'
WHERE Pk_IdVenta='${next.Pk_IdVenta}';`;
            }
        }

        // Validación de la fila actual, si no fue afectada por ajustes anteriores
        if (Math.abs(current.TotalVolumen - current.CantidadTotal) > 0.0001) {
            script += `
UPDATE [EstacionNSX].[dbo].[Venta]
SET CantidadTotal='${current.TotalVolumen.toFixed(2).replace(",", ".")}'
WHERE Pk_IdVenta='${current.Pk_IdVenta}';`;
        }

        if (Math.abs(current.TotalDinero - current.ValorTotal) > 0.0001) {
            script += `
UPDATE [EstacionNSX].[dbo].[Venta]
SET ValorTotal='${Math.round(current.TotalDinero)}'
WHERE Pk_IdVenta='${current.Pk_IdVenta}';`;
        }
    }

    document.getElementById("sqlOutput").innerText = script.trim();
    document.getElementById("sqlCard").classList.remove("d-none");
}

