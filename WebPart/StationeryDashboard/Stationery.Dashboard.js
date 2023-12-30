/// <reference path="../../JSDependency/BaseModule.js" />
var stationeryStockData = [];
var stationeryDetailsData = [];
var pendingApprovalData = [];

const gridOptions = {
    columnDefs: [
        { headerName: "Requisition No", field: 'RequisitionNo', maxWidth: 160, enableRowGroup: true },
        { headerName: "User Name", field: 'UserName', maxWidth: 200, enableRowGroup: true },
        { headerName: "User Id", field: 'UserId', maxWidth: 140 },
        { headerName: "Status", field: 'Status', maxWidth: 140, enableRowGroup: true },
        { headerName: "Pending With", field: 'PendingWith', maxWidth: 200, enableRowGroup: true },
        { headerName: "Required Item Name", field: 'RequiredItemName', maxWidth: 200, enableRowGroup: true },
        { headerName: "Required Quantity", field: 'RequiredQuantity', maxWidth: 200 },
        { headerName: "Present Stock", field: 'PresentStock', maxWidth: 140 },
        { headerName: "Issue Quantity", field: 'IssueQuantity', maxWidth: 140 },
        { headerName: "Unit Price", field: 'UnitPrice', maxWidth: 140 },
        { headerName: "Total Price", field: 'TotalPrice', maxWidth: 140 },
        { headerName: "Link", field: 'RequestLink', cellRenderer: viewActionCellRenderer, maxWidth: 100 },
        { headerName: "Created", field: 'Created', enableRowGroup: true, sort: 'desc' },
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
    gridOptions.api.setRowData(await MapMasterData());

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

const FetchStationeryStock = async () => {
    const query = `$select=Id,Title,MaterialCode,MaterialName,Unit,VendorName,UnitPrice,OpeningQuantity,IssuedQuantity,StockInHand,Date`;
    try {
        stationeryStockData = await GetByList('StationeryStock', query);
    } catch (error) {
        console.error(error);
    }
};

const FetchPendingApproval = async () => {
    const query = `$select=Id,Title,PendingWith/Title,PendingWith/Id,ProcessName,Status,RequestedBy/Title,RequestedBy/Id,RequestLink,Created&$expand=PendingWith,RequestedBy`;
    try {
        pendingApprovalData = await GetByList('PendingApproval', query);
    } catch (error) {
        console.error(error);
    }
};

const FetchStationeryDetails = async () => {
    const query = `$select=Id,Title,RequestedAmount,IssuedAmount,StationeryId,StationeryStockId,Unit,IssuedQuantity,RequestedQuantity,Author/Title,Author/Id,Created&$expand=Author`;
    try {
        stationeryDetailsData = await GetByList('StationeryDetails', query);
    } catch (error) {
        console.error(error);
    }
};


const MapMasterData = async () => {
    await FetchStationeryStock();
    await FetchPendingApproval();
    await FetchStationeryDetails();

    const mapStockById = id => stationeryStockData.find(item => item.ID === id);
    const masterDataArray = [];
    pendingApprovalData.forEach(pendingItem => {
        const detailsItems = stationeryDetailsData.filter(detailsItem => detailsItem.StationeryId === pendingItem.ID);

        detailsItems.forEach(detailsItem => {
            const stockItem = mapStockById(detailsItem.StationeryStockId);
            masterDataArray.push({
                "Author": pendingItem.Author,
                "AuthorId": pendingItem.AuthorId,
                "Title": pendingItem.Title,
                "Status": pendingItem.Status,
                "RequestLink": pendingItem.RequestLink,
                "PendingWith": pendingItem.PendingWith?.Title || 'N/A',
                "RequisitionNo": pendingItem.Title,
                "UserName": pendingItem.RequestedBy?.Title,
                "UserId": pendingItem.RequestedBy?.Id,
                "RequiredItemName": stockItem?.MaterialName || 'N/A',
                "RequiredQuantity": detailsItem?.RequestedQuantity || 'N/A',
                "PresentStock": stockItem?.StockInHand || 'N/A',
                "IssueQuantity": detailsItem?.IssuedQuantity || 'N/A',
                "UnitPrice": stockItem?.UnitPrice || 'N/A',
                "TotalPrice": detailsItem?.RequestedQuantity * stockItem?.UnitPrice || 'N/A',
                "Created": new Date(pendingItem.Created)
            });
        });
    });
    return masterDataArray;
};

