/// <reference path="../../JSDependency/BaseModule.js" />

const StationeryApp = angular.module('StationeryApp', [])

const ApprovalStatus = {
    Submitted: 'Submitted',
    ChangeRequested: 'Change Requested',
    
    Updated: 'Updated',
    LineManagerApproved: 'LM Approved',
    StationeryInChargeApproved: 'SIC Approved',
    Closed: 'Closed',
    Rejected: 'Rejected',
}
const ApprovalChain = {
    LineManager: null,
    StationeryInCharge: null,
}
const userEmail = _spPageContextInfo.userEmail;
const PendingApprovalUniqueId = new URLSearchParams(window.location.search).get('StationeryId');
const dev = new URLSearchParams(window.location.search).get('dev') ? true : false;
let IsEditable = false;
let ReadOnly = false;
let ChangeMode = false;
let CurerntPendingWithId = null;
let NewPendingWithId = null;
let CurrentStatus = null;
let StatusOnAction = null;
let StationeryId = null;
let PendingApprovalId = null;
let RequestedById = null;
let OnBehalf = null;
/**
 * @controller UserInfoController
 */
StationeryApp.controller('UserInfoController', function ($scope) {
    $scope.loading = true;

    const fetchUserData = async () => {
        try {
            let RequesterId = userId;
            if (PendingApprovalUniqueId) { /*  If The link is from PendingApproval */
                const query = `$select=Author\Id&expend=Author&$filter=substringof('${PendingApprovalUniqueId}',RequestLink)`;
                const data = await GetByList('PendingApproval', query);
                RequesterId = data[0].AuthorId;
            }
            const userdata = await GetByList("Employees", `$select=name,email/EMail,email/Title,deptId,phone,designation,LineManagerName,LineManagerEmail\Id&expend=LineManagerEmail&$expand=email,email/Id&$filter=email eq ${RequesterId}`);
            $scope.$apply(() => {
                $scope.userInfo = userdata[0];
                $scope.userInfo.name = userdata[0].email.Title;
                $scope.userInfo.email = userdata[0].email.EMail;
                $scope.loading = false;
            });
            await fetchApprvalChain();
        } catch (error) {
            console.error(error);
        }
    };
    const fetchApprvalChain = async () => {
        try {
            const query = `$select=Approver1Id&$filter=ProcessName eq 'Stationery' and DepertmentCode eq '${$scope.userInfo.deptId}'`
            const approvarinfo = await GetByList('ApproverInformation', query);
            ApprovalChain.LineManager = $scope.userInfo.LineManagerEmailId;
            ApprovalChain.StationeryInCharge = approvarinfo[0].Approver1Id;
        }
        catch (error) {
            console.error(error);
        }
    };

    fetchUserData();
});

/**
 * @controller MainController
 */
StationeryApp.controller('MainController', function ($scope, $http) {
    const vm = this;
    const ABS_BASE_URL = ABS_URL.replace('/sites/bprocess', '')
    vm.stationaryList = [];
    vm.inventory = [];


    const fetchOnBehalfOf = async () => {
        try {
            const query = `$select=email/Id,email/Title&$expand=email,email/Id`;
            const data = await GetByList('Employees', query);
            $scope.$apply(() => {
                vm.onBehalfOf = data.map(item => {
                    return {
                        name: item.email.Title,
                        email: item.email.Id
                    };
                });
                vm.selectedOnBehalf = "";
            });
        } catch (error) {
            console.error(error);
        }
    };
    fetchOnBehalfOf();
    vm.onBehalfChange = async function () {
        const fetchonBehalf = await GetByList('Employees', `$select=email/EMail,deptId,phone,designation,LineManagerEmail/Title&$expand=email,email/Id,LineManagerEmail/Id&$filter=email eq '${vm.selectedOnBehalf}'`);
        $scope.$apply(() => {
            vm.onBehalfEmail = fetchonBehalf[0].email.EMail;
            vm.onBehalfLineManager = fetchonBehalf[0].LineManagerEmail.Title;
            vm.onBehalfDeptId = fetchonBehalf[0].deptId;
            vm.onBehalfPhone = fetchonBehalf[0].phone;
            vm.onBehalfDesignation = fetchonBehalf[0].designation;
        });
        OnBehalf = vm.selectedOnBehalf;
    }

    if (PendingApprovalUniqueId) { /*  If The link is from PendingApproval */
        ReadOnly = vm.ReadOnly = true;
        console.log('PendingApprovalUniqueId:', PendingApprovalUniqueId);
        const fetchStationaryRequestData = async () => {
            try {
                const query = `$select=Id,Title,PendingWithId,ProcessName,Status,RequestedById,RequestLink,Author\Id&expend=Author&$filter=substringof('${PendingApprovalUniqueId}',RequestLink)`; // TODO: need to fix fragment uuid bug
                const data = await GetByList('PendingApproval', query);
                console.log('PendingApproval data:', data);
                if (data.length > 0) {
                    PendingApprovalId = data[0].Id;
                    RequestedById = data[0].RequestedById;
                    if (data[0].Status === ApprovalStatus.ChangeRequested && data[0].PendingWithId === userId) {
                        ChangeMode = vm.ChangeMode = true;
                        vm.showCommentBox = true;
                        vm.dontShowActionBtn = true;
                    } else {
                        ChangeMode = vm.ChangeMode = false;
                    }
                    const CurerntPendingWithName = await GetByList('Employees', `$select=name&$filter=email eq ${data[0].PendingWithId}`);
                    vm.currentPendingWith = CurerntPendingWithName[0]?.name;
                    StationeryId = data[0].Title.split('-')[1];
                    const Stationery_data = await GetByList('Stationery', `$select=*,OnBehalf/Id,OnBehalf/Title&$expand=OnBehalf,OnBehalf/Id&$filter=Id eq ${StationeryId}`);
                    console.log('Stationery data:', Stationery_data);
                    const StationeryDetails_data = await GetByList('StationeryDetails', `$select=*&$filter=StationeryId eq ${StationeryId}`);
                    console.log('StationeryDetails data:', StationeryDetails_data);
                    const StationeryStock_data = await GetByList('StationeryStock', `$select=Id,ContentTypeId,Title,ComplianceAssetId,MaterialCode,MaterialName,Unit,VendorName,UnitPrice,OpeningQuantity,IssuedQuantity,StockInHand,Date,ID,AttachmentFiles&$expand=AttachmentFiles`);
                    console.log('StationeryStock data:', StationeryStock_data);
                    const StationeryLog_data = await GetByList('StationeryLog', `$select=Author/Title,Comment,Created,Status&$expand=Author/Id&$filter=StationeryId eq ${StationeryId}`);
                    console.log('StationeryLog data:', StationeryLog_data);
                    const attachments = await GetByList('StationeryAttachment', `$select=Author/Title,Created,AttachmentFiles&$expand=AttachmentFiles,Author/Id&$filter=StationeryId eq ${StationeryId}`);
                    console.log('attachments:', attachments);

                    vm.selectedOnBehalf = Stationery_data[0].OnBehalfId;
                    vm.selectedOnBehalfName = Stationery_data[0].OnBehalf.Title;
                    vm.hideSelect2 = true;
                    vm.onBehalfChange();

                    vm.requestHistory = StationeryLog_data.map(item => {
                        return {
                            Comment: item.Comment,
                            Date: item.Created,
                            Status: item.Status,
                            Author: item.Author.Title
                        };
                    });

                    vm.attachments = attachments.map(item => {
                        return {
                            Name: item.AttachmentFiles.results[0].FileName,
                            Url: ABS_BASE_URL + item.AttachmentFiles.results[0].ServerRelativePath.DecodedUrl,
                            Author: item.Author.Title,
                            Date: item.Created
                        };
                    });

                    // Render StationeryRequest data //
                    $scope.$apply(() => {
                        vm.inventory = StationeryDetails_data.map(item => {
                            const selectedItem = StationeryStock_data.find(stock => stock.Id === item.StationeryStockId);
                            return {
                                selectedItem: selectedItem,
                                quantity: item.RequestedQuantity,
                                TotalIssuedQuantity: item.IssuedQuantity,
                                imageUrl: ABS_BASE_URL + selectedItem.AttachmentFiles.results[0].ServerRelativePath.DecodedUrl
                            };
                        });
                        vm.selectedFloor = Stationery_data[0].Floor;
                        vm.inventory.forEach(item => {
                            item.selectedItem = vm.stationaryList.find(stationary => stationary.Id === item.selectedItem.Id);
                        });
                    });
                    CurerntPendingWithId = data[0].PendingWithId;
                    CurrentStatus = vm.currentStatus = data[0].Status;

                    /* Button Config */
                    if (CurerntPendingWithId === userId && CurerntPendingWithId !== RequestedById || dev) {
                        vm.showApproveBtn = true;
                        vm.showCommentBox = true;
                    }
                    if (CurerntPendingWithId === userId && CurerntPendingWithId === RequestedById && CurrentStatus === ApprovalStatus.StationeryInChargeApproved || dev) {
                        vm.showCloseBtn = true;
                        vm.showCommentBox = true;
                    }

                    /* ------STRT WORKFLOW------ */
                    /* Submitted */
                    if (CurrentStatus === ApprovalStatus.Submitted && CurerntPendingWithId === ApprovalChain.LineManager || CurrentStatus === ApprovalStatus.Updated && CurerntPendingWithId === ApprovalChain.LineManager) {
                        NewPendingWithId = ApprovalChain.StationeryInCharge;
                        StatusOnAction = ApprovalStatus.LineManagerApproved;
                    }/*  Line Manager Approved */
                    else if (CurrentStatus === ApprovalStatus.LineManagerApproved && CurerntPendingWithId === ApprovalChain.StationeryInCharge || CurrentStatus === ApprovalStatus.Updated && CurerntPendingWithId === ApprovalChain.StationeryInCharge) {
                        /*  NewPendingWithId = RequestedById; */
                        NewPendingWithId = null;
                        StatusOnAction = ApprovalStatus.StationeryInChargeApproved;
                        if (CurerntPendingWithId === userId)
                            vm.showTotalIssuedQuantityInput = true;
                    }/* Close */
                    /* else if (CurrentStatus === ApprovalStatus.StationeryInChargeApproved && CurerntPendingWithId === userId) {
                        NewPendingWithId = null;
                        vm.showTotalIssuedQuantityText = true;
                        StatusOnAction = ApprovalStatus.Closed;
                    } */ /* Change Request */
                    else if (CurrentStatus === ApprovalStatus.ChangeRequested && CurerntPendingWithId === userId) {
                        StatusOnAction = ApprovalStatus.Submitted;
                        const fetchLastPendingWith = async () => {
                            try {
                                const query = `$select=Author\Id&$filter=Title eq 'S-${StationeryId}'&$orderby=Created desc`;
                                const data = await GetByList('StationeryLog', query);
                                return data[0].AuthorId;
                            } catch (error) {
                                console.error('Error fetching StationeryLog data:', error);
                            }
                        }
                        NewPendingWithId = await fetchLastPendingWith();
                    }
                    /* ------END WORKFLOW------ */
                    if (CurrentStatus === ApprovalStatus.StationeryInChargeApproved || CurrentStatus === ApprovalStatus.Closed)
                        vm.showTotalIssuedQuantityText = true;

                    let NewPendingWithIdName = await GetByList('Employees', `$select=name&$filter=email eq ${NewPendingWithId}`);
                    $scope.$apply(() => {
                        vm.nextPendingWith = NewPendingWithIdName[0]?.name;
                    });
                }
            } catch (error) {
                console.error('Error fetching StationeryRequest data:', error);
            }
        };
        fetchStationaryRequestData();
    }
    else {
        vm.showCommentBox = true;
    }

    const fetchStationaryData = async () => {
        try {
            const query = "$select=Id,ContentTypeId,Title,ComplianceAssetId,MaterialCode,MaterialName,Unit,VendorName,UnitPrice,OpeningQuantity,IssuedQuantity,StockInHand,Date,ID,AttachmentFiles&$expand=AttachmentFiles";
            const data = await GetByList('StationeryStock', query);
            vm.stationaryList = data;
            $scope.$apply(() => {
                vm.inventory.push({
                    selectedItem: vm.stationaryList[0],
                    quantity: 0,
                    imageUrl: ABS_BASE_URL + vm.stationaryList[0].AttachmentFiles.results[0].ServerRelativePath.DecodedUrl
                });
            });
        } catch (error) {
            console.error('Error fetching StationeryStock data:', error);
        }
    };

    fetchStationaryData();

    vm.addRow = function () {
        vm.inventory.push({
            selectedItem: vm.stationaryList[0],
            quantity: 0,
            imageUrl: ABS_BASE_URL + vm.stationaryList[0].AttachmentFiles.results[0].ServerRelativePath.DecodedUrl
        });
    };

    vm.updateImage = function (item) {
        item.imageUrl = ABS_BASE_URL + item.selectedItem.AttachmentFiles.results[0].ServerRelativePath.DecodedUrl;
    };

    vm.removeRow = function (index) {
        if (vm.inventory.length === 1) return;
        vm.inventory.splice(index, 1);
    };

    vm.submitRequest = function () {
        const TotalRequestedAmount = vm.inventory.reduce((total, item) => {
            return total + (item.selectedItem.UnitPrice * item.quantity);
        }, 0);
        const TotalRequestedQuantity = vm.inventory.reduce((total, item) => {
            return total + item.quantity;
        }, 0);

        const Stationery_data = {
            "__metadata": { "type": "SP.Data.StationeryListItem" },
            "Title": "S",
            "PendingWithId": ApprovalChain.LineManager, // 1st approver
            "TotalRequestedAmount": TotalRequestedAmount,
            "TotalRequestedQuantity": TotalRequestedQuantity,
            "Status": ApprovalStatus.Submitted,
            "Floor": vm.selectedFloor,
            "OnBehalfId": OnBehalf,
        };

        vm.loading = true;
        $('#SuccessModal').modal('show');


        const addStationeryRequestItems = async (StationeryId) => {
            try {
                const promises = vm.inventory.map(item => {
                    const StationaryDetails_data = {
                        "__metadata": { "type": "SP.Data.StationaryDetailsListItem" },
                        "Title": `S-${StationeryId}`,
                        "RequestedAmount": item.selectedItem.UnitPrice * item.quantity,
                        "IssuedAmount": 0,
                        "StationeryId": StationeryId,
                        "StationeryStockId": item.selectedItem.Id,
                        "Unit": item.selectedItem.Unit,
                        "IssuedQuantity": 0,
                        "RequestedQuantity": item.quantity,
                    };
                    return AddListItem("StationeryDetails", StationaryDetails_data);
                });
                await Promise.all(promises);
            } catch (error) {
                console.error('Error adding StationeryRequestItems:', error);
            }
        }

        const addPendingApproval = async (StationeryId) => {
            try {
                const PendingApproval_data = {
                    "__metadata": { "type": "SP.Data.PendingApprovalListItem" },
                    "Title": `S-${StationeryId}`,
                    "PendingWithId": ApprovalChain.LineManager, // 1st approver
                    "RequestLink": `${ABS_URL}/SitePages/stationery.aspx?StationeryId=${uuid.v4()}`,
                    "RequestedById": userId,
                    "ProcessName": "Stationery",
                    "Status": ApprovalStatus.Submitted,
                };
                await AddListItem("PendingApproval", PendingApproval_data);
            } catch (error) {
                console.error('Error adding PendingApproval:', error);
            }
        };

        const addStationeryRequest = async () => {
            try {
                const res = await AddListItem("Stationery", Stationery_data);
                await addStationeryRequestItems(res.Id);
                await addPendingApproval(res.Id);
                await addHistory(res.Id, ApprovalStatus.Submitted);
                await SaveAllAttachments(res.Id);

                $scope.$apply(() => {
                    vm.loading = false;
                    vm.showMCloseBtn = false;
                    vm.ModalMessage = `Your request has been submitted successfully. ID #${res.Id}`;
                    console.log('StationeryRequest added successfully:', res);
                });
            } catch (error) {
                $scope.$apply(() => {
                    vm.loading = false;
                    vm.showMCloseBtn = true;
                    vm.ModalMessage = `An error occurred while submitting your request. Please try again later.`;
                    console.error('Error adding StationeryRequest:', error);
                });
            }
        }

        addStationeryRequest();
    };
    const addHistory = async (StationeryId, Status) => {
        try {
            const History_data = {
                "__metadata": { "type": "SP.Data.StationaryLogListItem" },
                "Title": `S-${StationeryId}`,
                "StationeryId": StationeryId,
                "Status": Status,
                "Comment": vm.comment
            };
            await AddListItem("StationeryLog", History_data);
        } catch (error) {
            console.error('Error adding History:', error);
        }
    };
    const AddAttachment = async (file, RequestId) => {
        const ListName = "StationeryAttachment";
        const data = {
            'Title': `ST-${RequestId}`,
            'StationeryId': RequestId,
            '__metadata': { "type": "SP.Data.StationaryAttachmentListItem" },
        }
        try {
            const response = await AddListItem(ListName, data);
            await uploadFileToSharePoint(response.ID, file, ListName);
        } catch (error) {
            console.log(error);
        }
    }
    const SaveAllAttachments = async (id) => {
        let fileInputs = $("#attachFilesContainer input:file");
        const filesToUpload = Array.from(fileInputs)
            .map((input) => input.files[0])
            .filter((file) => file);
        return new Promise((resolve, reject) => {
            Promise.all(filesToUpload.map((file) => AddAttachment(file, id)))
                .then(() => {
                    resolve();
                })
                .catch((error) => {
                    $scope.loading = false;
                    reject('Uploaded is false.');
                    console.error('Error uploading files:', error);
                    reject(error);
                });
        });
    }
    vm.updateStatus = function (status) { // TODO : need to add Change Request functionality
        const updateStationeryRequest = async () => {
            try {
                vm.loading = true;
                $('#SuccessModal').modal('show');
                if (status === ApprovalStatus.Rejected) {
                    StatusOnAction = ApprovalStatus.Rejected;
                    NewPendingWithId = null;
                }
                if (status === ApprovalStatus.ChangeRequested) {
                    StatusOnAction = ApprovalStatus.ChangeRequested;
                    NewPendingWithId = RequestedById;
                }

                let Stationery_data = {
                    "__metadata": { "type": "SP.Data.StationeryListItem" },
                    "PendingWithId": NewPendingWithId,
                    "Status": StatusOnAction
                };

                if (CurerntPendingWithId === ApprovalChain.StationeryInCharge) {
                    const TotalApprovedAmount = vm.inventory.reduce((total, item) => {
                        return total + (item.selectedItem.UnitPrice * item.TotalIssuedQuantity);
                    }, 0);
                    const TotalIssuedQuantity = vm.inventory.reduce((total, item) => {
                        return total + item.TotalIssuedQuantity;
                    }, 0);

                    Stationery_data = {
                        ...Stationery_data,
                        "TotalApprovedAmount": TotalApprovedAmount,
                        "TotalIssuedQuantity": TotalIssuedQuantity,
                    }

                    const StationeryDetails_data = vm.inventory.map(item => {
                        return {
                            "__metadata": { "type": "SP.Data.StationaryDetailsListItem" },
                            "Id": item.selectedItem.Id,
                            "Title": `S-${StationeryId}`,
                            "IssuedAmount": item.selectedItem.UnitPrice * item.TotalIssuedQuantity,
                            "IssuedQuantity": item.TotalIssuedQuantity,
                        };
                    });

                    const promises = StationeryDetails_data.map(async item => {
                        const query = `$select=Id&$filter=StationeryId eq ${StationeryId} and StationeryStockId eq ${item.Id}`;
                        const data = await GetByList('StationeryDetails', query);
                        return UpdateListItem("StationeryDetails", data[0].Id, item);
                    });
                    await Promise.all(promises);

                    const StationeryStock_data = vm.inventory.map(item => {
                        return {
                            "__metadata": { "type": "SP.Data.StationaryStockListItem" },
                            "Id": item.selectedItem.Id,
                            "IssuedQuantity": item.selectedItem.IssuedQuantity + item.TotalIssuedQuantity,
                            "StockInHand": item.selectedItem.StockInHand - item.TotalIssuedQuantity,
                        };
                    });

                    const ReduceStockPromises = StationeryStock_data.map(item => {
                        return UpdateListItem("StationeryStock", item.Id, item);
                    });
                    await Promise.all(ReduceStockPromises);
                }

                const PendingApproval_data = {
                    "__metadata": { "type": "SP.Data.PendingApprovalListItem" },
                    "PendingWithId": NewPendingWithId,
                    "Status": StatusOnAction,
                };
                await UpdateListItem("Stationery", StationeryId, Stationery_data);
                await UpdateListItem("PendingApproval", PendingApprovalId, PendingApproval_data);
                await addHistory(StationeryId, StatusOnAction);
                await SaveAllAttachments(StationeryId);
                $scope.$apply(() => {
                    vm.loading = false;
                    vm.showMCloseBtn = false;
                    vm.ModalMessage = `Request has been ${StatusOnAction}.`;
                });
            } catch (error) {
                $scope.$apply(() => {
                    vm.loading = false;
                    vm.showMCloseBtn = true;
                    vm.ModalMessage = `An error occurred while submitting your request. Please try again later.`;
                    console.error('Error adding StationeryRequest:', error);
                });
            }
        }
        updateStationeryRequest();
    };
    vm.updateRequest = function () {
        const updateStationeryRequest = async () => {
            try {
                vm.loading = true;
                $('#SuccessModal').modal('show');
                const TotalRequestedAmount = vm.inventory.reduce((total, item) => {
                    return total + (item.selectedItem.UnitPrice * item.quantity);
                }, 0);
                const TotalRequestedQuantity = vm.inventory.reduce((total, item) => {
                    return total + item.quantity;
                }, 0);

                /* StationeryDetails dm */
                const StationeryDetails_data = vm.inventory.map(item => {
                    return {
                        "__metadata": { "type": "SP.Data.StationaryDetailsListItem" },
                        "Id": item.selectedItem.Id,
                        "Title": `S-${StationeryId}`,
                        "RequestedAmount": item.selectedItem.UnitPrice * item.quantity,
                        "RequestedQuantity": item.quantity,
                    };
                });

                const promises = StationeryDetails_data.map(async item => {
                    const query = `$select=Id&$filter=StationeryId eq ${StationeryId} and StationeryStockId eq ${item.Id}`;
                    const data = await GetByList('StationeryDetails', query);
                    return UpdateListItem("StationeryDetails", data[0].Id, item);
                });
                await Promise.all(promises);

                /* Stationery dm */
                const Stationery_data = {
                    "__metadata": { "type": "SP.Data.StationeryListItem" },
                    "TotalRequestedAmount": TotalRequestedAmount,
                    "TotalRequestedQuantity": TotalRequestedQuantity,
                    "Floor": vm.selectedFloor,
                    "Status": ApprovalStatus.Updated,
                    "PendingWithId": NewPendingWithId,
                };
                await UpdateListItem("Stationery", StationeryId, Stationery_data);

                /* PendingApproval dm */
                const PendingApproval_data = {
                    "__metadata": { "type": "SP.Data.PendingApprovalListItem" },
                    "PendingWithId": NewPendingWithId,
                    "Status": ApprovalStatus.Updated,
                };
                await UpdateListItem("PendingApproval", PendingApprovalId, PendingApproval_data);

                await addHistory(StationeryId, ApprovalStatus.Updated);
                $scope.$apply(() => {
                    vm.loading = false;
                    vm.showMCloseBtn = false;
                    vm.ModalMessage = `Your request has been updated successfully.`;
                });
            } catch (error) {
                $scope.$apply(() => {
                    vm.loading = false;
                    vm.showMCloseBtn = true;
                    vm.ModalMessage = `An error occurred while submitting your request. Please try again later.`;
                    console.error('Error adding StationeryRequest:', error);
                });
            }
        }
        updateStationeryRequest();
    }
});
