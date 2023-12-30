/// <reference path="../../JSDependency/BaseModule.js" />

const gridOptions = {
    columnDefs: [
        { headerName: "Asset Subcategory", field: "subCategory", enableRowGroup: true },
        { headerName: "Model", field: "model" },
        { headerName: "Asset Number", field: "assetNumber" },
        { headerName: "Acquisition Type", field: "acquisitionType", enableRowGroup: true },
        { headerName: "Vendor", field: "purchaseVendor" },
        { headerName: "Asset Title", field: "assetTitle" },
        { headerName: "Employee ID", field: "employeeId" },
        { headerName: "Asset Users", field: "AssetUsers.Title", enableRowGroup: true },
        { headerName: "Email", field: "email" },
        { headerName: "Purchase Date", field: "purchaseDate" },
        { headerName: "Manufacturer", field: "manufacturer" },
        { headerName: "Product SL No", field: "productSLNo" },
        { headerName: "Warranty Period (To)", field: "warrantyPeriodTo" },
    ],
    defaultColDef: {
        sortable: true,
        resizable: true,
    },
    animateRows: true,
    rowGroupPanelShow: 'always',
    flex: 1,
    minWidth: 100,

    paginationPageSize: 100,
    suppressRowClickSelection: true,
    groupSelectsChildren: true,
    rowSelection: 'multiple',
    rowGroupPanelShow: 'always',
    pivotPanelShow: 'always',
    pagination: true,
};


$(document).ready(async function () {
    new agGrid.Grid($('#dataGrid')[0], gridOptions);
    gridOptions.api.setRowData(await FetchData());

    const filterText = $('#filter-text-box');
    filterText.on('input', function () {
        gridOptions.api.setQuickFilter($(this).val());
    });
    filterText.on('keydown', function (event) {
        if (event.keyCode === 13 /* Enter */) {
            event.preventDefault();
        }
    });
});

const onBtExport = () => gridOptions.api.exportDataAsExcel();

function viewActionCellRenderer(params) {
    if (params.value === undefined) return null;
    return LinkRenderer(params, 'View');
}

function LinkRenderer(params, label) {
    const viewActionValue = params.value;

    return $('<a>', {
        href: viewActionValue,
        text: label,
        click: function (event) {
            event.preventDefault();
            window.open(viewActionValue, '_blank');
        }
    })[0]
}

const FetchData = async () => {
    const query = `$select=assetNumber,subCategory,acquisitionType,purchaseVendor,model,assetTitle,employeeId,AssetUsers/Title,email,purchaseDate,manufacturer,productSLNo,warrantyPeriodTo&$expand=AssetUsers`;

    try {
        const data = await GetByList('ItAssetMaster', query);
        return data.map(item => ({
            ...item,
            warrantyPeriodTo: new Date(item.warrantyPeriodTo),
            purchaseDate: new Date(item.purchaseDate)
        }));
    } catch (error) {
        console.error(error);
    }
};


