let data = [] // Almacenar los datos originales del Excel

// Función para leer el archivo Excel
const mostrarExcel = async () => {
    const archivoExcel = 'test.xlsx' // Nombre del archivo Excel

    // Cargar la librería SheetJS
    await new Promise((resolve, reject) => {
        const script = document.createElement('script')
        script.src = 'https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js'
        script.onload = resolve
        script.onerror = reject
        document.head.appendChild(script)
    })

    // Leer el archivo Excel
    const response = await fetch(archivoExcel)
    const arrayBuffer = await response.arrayBuffer()
    const workbook = XLSX.read(arrayBuffer, { type: 'array' })

    // Obtener la primera hoja de cálculo
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]

    // Convertir los datos de la hoja en JSON
    data = XLSX.utils.sheet_to_json(sheet, { header: 1 })

    // Mostrar los datos en una tabla HTML
    const table = document.getElementById('excel-table')
    data.forEach(row => {
        const newRow = table.insertRow()
        row.forEach(cell => {
            const newCell = newRow.insertCell()
            newCell.textContent = cell
        })
    })

    // Agregar opciones al campo de selección para filtrar por columna
    const filterCol = document.getElementById('filterCol')
    data[0].forEach((column, index) => {
        const option = document.createElement('option')
        option.value = index
        option.textContent = column
        filterCol.appendChild(option)
    })
}

// Función para filtrar la tabla por columna
const filtrarTabla = () => {
    const filterColIndex = document.getElementById('filterCol').value
    const filterText = document.getElementById('filterText').value.toLowerCase()

    const table = document.getElementById('excel-table')
    table.innerHTML = '' // Limpiar la tabla

    data.forEach((row, rowIndex) => {
        if (rowIndex === 0) {
            // Agregar la fila de encabezados
            const headerRow = table.insertRow()
            row.forEach(cell => {
                const headerCell = headerRow.insertCell()
                headerCell.textContent = cell
            })
        } else {
            if (row[filterColIndex].toString().toLowerCase().includes(filterText)) {
                // Agregar filas que cumplan con el filtro
                const newRow = table.insertRow()
                row.forEach(cell => {
                    const newCell = newRow.insertCell()
                    newCell.textContent = cell
                })
            }
        }
    })
}

// Llamar a la función para mostrar el Excel al cargar la página
mostrarExcel()