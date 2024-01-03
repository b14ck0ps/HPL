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
        { headerName: "Manufacturer", field: "manufacturer" },
        { headerName: "Product SL No", field: "productSLNo" },
        { headerName: "Warranty Period (To)", field: "warrantyPeriodTo" },
        { headerName: "Purchase Date", field: "purchaseDate" },
        { headerName: "Asset Actual Cost", field: "purchasePrice" },
        { headerName: "Asset Remaining Cost", field: "RemainingCost" },
        { headerName: "Asset EMI", field: "EMI" },
        { headerName: "Link", field: 'Id', cellRenderer: viewActionCellRenderer, maxWidth: 100 },
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
    await GetAssetLifecycle();
    gridOptions.api.setRowData(calculateRemainingCostAndEMI(await FetchData(), await GetAssetLifecycle()));

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

    const href = `${ABS_URL}/SitePages/ItAssetForm.aspx?AssetId=${viewActionValue}`;
    return $('<a>', {
        href: href,
        text: label,
        click: function (event) {
            event.preventDefault();
            window.open(href, '_blank');
        }
    })[0]
}

const FetchData = async () => {
    const query = `$select=Id,assetNumber,subCategory,acquisitionType,purchaseVendor,model,assetTitle,employeeId,AssetUsers/Title,email,purchaseDate,manufacturer,productSLNo,warrantyPeriodTo,purchasePrice&$expand=AssetUsers&$filter=acquisitionType eq 'Damage' or acquisitionType eq 'Lost'`;

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

const GetAssetLifecycle = async () => {
    const query = `$select=Title,LostWithin,AcquisitionCost,RealizationPeriod`;

    try {
        return await GetByList('AssetLifecycle', query);

    } catch (error) {
        console.error(error);
    }
}

function calculateRemainingCostAndEMI(data, AssetLifecycle) {
    return data.map((item) => {
        const purchaseDate = new Date(item.purchaseDate);
        const currentDate = new Date();

        // Calculate age of purchase in months
        const ageInMonths = (currentDate.getFullYear() - purchaseDate.getFullYear()) * 12 +
            currentDate.getMonth() - purchaseDate.getMonth();

        // Find realizationPeriod based on age in months
        const realizationPeriodData = AssetLifecycle.find(
            (lifecycle) => ageInMonths <= lifecycle.LostWithin
        );

        if (realizationPeriodData) {
            const remainingCost = item.purchasePrice * realizationPeriodData.AcquisitionCost;
            const EMI = remainingCost / realizationPeriodData.RealizationPeriod;

            return {
                ...item,
                RemainingCost: remainingCost.toFixed(2),
                EMI: EMI.toFixed(2),
            };
        } else {
            return {
                ...item,
                RemainingCost: "0.00",
                EMI: "0.00",
            };
        }
    });
}




