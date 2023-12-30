/// <reference path="../../JSDependency/BaseModule.js" />

const StationeryForcast = angular.module('StationeryForcast', [])

StationeryForcast.controller('ForcastController', function ($scope) {
    const vm = this;
    let StationeryDetails = [];
    let StationeryStock = [];

    const fetchData = async () => {
        try {

            const currentDate = moment();
            const lastThreeMonths = currentDate.clone().subtract(3, 'months').format('YYYY-MM-DDTHH:mm:ss') + 'Z';

            const StationeryDetailsQuery = `$select=Id,RequestedAmount,IssuedAmount,StationeryStockId,Unit,IssuedQuantity,RequestedQuantity,Created&$filter=Created ge '${lastThreeMonths}'`;

            const StationeryStockQuery = `$select=Id,MaterialName,IssuedQuantity,StockInHand,Date`;

            StationeryDetails = await GetByList('StationeryDetails', StationeryDetailsQuery);
            StationeryStock = await GetByList('StationeryStock', StationeryStockQuery);
            const mappedData = StationeryDetails.map(detail => {
                const stock = StationeryStock.find(stock => stock.Id === detail.StationeryStockId);
                return {
                    ItemName: stock ? stock.MaterialName : '',
                    RequestedAmount: detail.RequestedAmount,
                    StockInHand: stock ? stock.StockInHand : 0,
                    Created: detail.Created,
                };
            });

            //Group by ItemName
            const groupedData = mappedData.reduce((acc, item) => {
                if (!acc[item.ItemName]) {
                    acc[item.ItemName] = [];
                }
                acc[item.ItemName].push(item);
                return acc;
            }, {});

            //Calculate requested amount for the last 3 months
            const result = Object.entries(groupedData).map(([itemName, details]) => {
                const monthlyAmounts = [0, 0, 0];

                details.forEach(detail => {
                    const createdDate = moment(detail.Created);
                    const monthsAgo = currentDate.diff(createdDate, 'months');

                    if (monthsAgo >= 0 && monthsAgo < 3) {
                        monthlyAmounts[monthsAgo] += detail.RequestedAmount;
                    }
                });

                return {
                    ItemName: itemName.toLowerCase(),
                    Month3: monthlyAmounts[0],
                    Month2: monthlyAmounts[1],
                    Month1: monthlyAmounts[2],
                    StockInHand: details[0].StockInHand,
                };
            });

            // Group the result by ItemName again to ensure correct grouping
            const MappedByMonths = result.reduce((acc, item) => {
                const key = item.ItemName;
                if (!acc[key]) {
                    acc[key] = {
                        ItemName: item.ItemName,
                        Month3: 0,
                        Month2: 0,
                        Month1: 0,
                        StockInHand: 0,
                    };
                }
                acc[key].Month3 += item.Month3;
                acc[key].Month2 += item.Month2;
                acc[key].Month1 += item.Month1;
                acc[key].StockInHand = item.StockInHand;
                return acc;
            }, []);

            // Convert the grouped result back to an array
            const MappedByMonthsArray = Object.values(MappedByMonths);
            $scope.$apply(() => {
                console.log(StationeryDetails);
                console.log(StationeryStock);
                console.log(MappedByMonthsArray);
                vm.StationeryForcastData = MappedByMonthsArray.map((item) => {
                    return {
                        ...item,
                        TotalUsageIn3Month: item.Month3 + item.Month2 + item.Month1,
                        ThreeMonthsAverageUsage: (item.Month3 + item.Month2 + item.Month1) / 3,
                        RequirementOf2Month: ((item.Month3 + item.Month2 + item.Month1) / 3) * 2,
                        NextMonthRequirement: (((item.Month3 + item.Month2 + item.Month1) / 3) * 2) - item.StockInHand,
                    }
                });
                console.log(vm.StationeryForcastData);
                $scope.loading = false;
            });
        } catch (error) {
            console.error(error);
        }
    };

    vm.exportToExcel = function () {
        var data = vm.StationeryForcastData;

        if (data.length === 0) {
            alert('No data to export.');
            return;
        }

        var dataArray = data.map(function (item) {
            return [
                item.ItemName,
                item.Month1,
                item.Month2,
                item.Month3,
                item.StockInHand,
                item.TotalUsageIn3Month,
                item.ThreeMonthsAverageUsage,
                item.RequirementOf2Month,
                item.NextMonthRequirement
            ];
        });

        var headers = [
            "ItemName",
            "ThreeMonthsAgo",
            "TwoMonthsAgo",
            "PreviousMonth",
            "StockInHand",
            "TotalUsageIn3Month",
            "ThreeMonthsAverageUsage",
            "RequirementOf2Month",
            "NextMonthRequirement"
        ];
        dataArray.unshift(headers);

        var ws = XLSX.utils.aoa_to_sheet(dataArray);

        var wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet 1');

        XLSX.writeFile(wb, 'StationeryForcast.xlsx');
    };

    fetchData();
});