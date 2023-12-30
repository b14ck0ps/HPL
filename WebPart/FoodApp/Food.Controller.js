/// <reference path="../../JSDependency/BaseModule.js" />

const FoodApp = angular.module('FoodApp', [])

const ApprovalStatus = {
    Submitted: 'Submitted',
    ChangeRequested: 'Change Requested',
    Updated: 'Updated',
    LineManagerApproved: 'LM Approved',
    FoodInChargeApproved: 'FIC Approved',
    Closed: 'Closed',
    Rejected: 'Rejected',
}
const ApprovalChain = {
    LineManager: null,
    FoodInCharge: null,
}
const userEmail = _spPageContextInfo.userEmail;
const PendingApprovalUniqueId = new URLSearchParams(window.location.search).get('FoodId');
const dev = new URLSearchParams(window.location.search).get('dev') ? true : false;
let IsEditable = false;
let ReadOnly = false;
let ChangeMode = false;
let CurerntPendingWithId = null;
let NewPendingWithId = null;
let CurrentStatus = null;
let StatusOnAction = null;
let FoodId = null;
let PendingApprovalId = null;
let RequestedById = null;
let OnBehalf = null;
/**
 * @controller UserInfoController
 */
FoodApp.controller('UserInfoController', function ($scope) {
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
            const query = `$select=Approver1Id&$filter=ProcessName eq 'Food' and DepertmentCode eq '${$scope.userInfo.deptId}'`
            const approvarinfo = await GetByList('ApproverInformation', query);
            ApprovalChain.LineManager = $scope.userInfo.LineManagerEmailId;
            ApprovalChain.FoodInCharge = approvarinfo[0].Approver1Id;
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
FoodApp.controller('MainController', function ($scope, $http) {
    const vm = this;
    const ABS_BASE_URL = ABS_URL.replace('/sites/bprocess', '')
    vm.FoodList = [];
   
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
        const fetchFoodRequestData = async () => {
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
                    FoodId = data[0].Title.split('-')[1];
                    const Food_data = await GetByList('Food', `$select=*,OnBehalf/Id,OnBehalf/Title&$expand=OnBehalf,OnBehalf/Id&$filter=Id eq ${FoodId}`);
                    console.log('Food data:', Food_data);
                    const FoodLog_data = await GetByList('FoodLog', `$select=Author/Title,Comment,Created,Status&$expand=Author/Id&$filter=FoodId eq ${FoodId}`);
                    console.log('FoodLog data:', FoodLog_data);
                    const attachments = await GetByList('FoodAttachment', `$select=Author/Title,Created,AttachmentFiles&$expand=AttachmentFiles,Author/Id&$filter=FoodId eq ${FoodId}`);
                    console.log('attachments:', attachments);

                    vm.selectedOnBehalf = Food_data[0].OnBehalfId;
                    vm.selectedOnBehalfName = Food_data[0].OnBehalf.Title;
                    vm.hideSelect2 = true;
                    vm.onBehalfChange();

                    vm.requestHistory = FoodLog_data.map(item => {
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

                    // rendering data
                    vm.agenda = Food_data[0].Agenda;
                    vm.participantsNumber = Food_data[0].ParticipantsNumber;
                    vm.plannedCost = Food_data[0].PlannedCost;
                    vm.meetingTime = Food_data[0].MeetingTime;
                    vm.selectedmeetingRoomName = Food_data[0].MeetingRoomName;
                    vm.meetingDate = Food_data[0].MeetingDate;
                    vm.selectedfoodCategory = Food_data[0].FoodCategory;
                    vm.selectedfoodSetMenu = Food_data[0].FoodSetMenu;
                    vm.costCenter = Food_data[0].CostCenter;
                    vm.gLNumber = Food_data[0].GLNumber;
                   
                    CurerntPendingWithId = data[0].PendingWithId;
                    CurrentStatus = vm.currentStatus = data[0].Status;

                    /* Button Config */
                    if (CurerntPendingWithId === userId && CurerntPendingWithId !== RequestedById || dev) {
                        vm.showApproveBtn = true;
                        vm.showCommentBox = true;
                    }
                    if (CurerntPendingWithId === userId && CurerntPendingWithId === RequestedById && CurrentStatus === ApprovalStatus.FoodInChargeApproved || dev) {
                        vm.showCloseBtn = true;
                        vm.showCommentBox = true;
                    }

                    /* ------STRT WORKFLOW------ */
                    /* Submitted */
                    if (CurrentStatus === ApprovalStatus.Submitted && CurerntPendingWithId === ApprovalChain.LineManager || CurrentStatus === ApprovalStatus.Updated && CurerntPendingWithId === ApprovalChain.LineManager) {
                        NewPendingWithId = ApprovalChain.FoodInCharge;
                        StatusOnAction = ApprovalStatus.LineManagerApproved;
                    }/*  Line Manager Approved */
                    else if (CurrentStatus === ApprovalStatus.LineManagerApproved && CurerntPendingWithId === ApprovalChain.FoodInCharge || CurrentStatus === ApprovalStatus.Updated && CurerntPendingWithId === ApprovalChain.FoodInCharge) {
                        /*  NewPendingWithId = RequestedById; */
                        NewPendingWithId = null;
                        StatusOnAction = ApprovalStatus.FoodInChargeApproved;
                       
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
                                const query = `$select=Author\Id&$filter=Title eq 'S-${FoodId}'&$orderby=Created desc`;
                                const data = await GetByList('FoodLog', query);
                                return data[0].AuthorId;
                            } catch (error) {
                                console.error('Error fetching FoodLog data:', error);
                            }
                        }
                        NewPendingWithId = await fetchLastPendingWith();
                    }
                    /* ------END WORKFLOW------ */
                    if (CurrentStatus === ApprovalStatus.FoodInChargeApproved || CurrentStatus === ApprovalStatus.Closed)
                        vm.showTotalIssuedQuantityText = true;

                    let NewPendingWithIdName = await GetByList('Employees', `$select=name&$filter=email eq ${NewPendingWithId}`);
                    $scope.$apply(() => {
                        vm.nextPendingWith = NewPendingWithIdName[0]?.name;
                    });
                }
            } catch (error) {
                console.error('Error fetching FoodRequest data:', error);
            }
        };
        fetchFoodRequestData();
    }
    else {
        vm.showCommentBox = true;
    }


    
    vm.submitRequest = function () {
        
        const Food_data = {
            "__metadata": { "type": "SP.Data.FoodListItem" },
            "Title": "F",
            "Agenda": vm.agenda,
            "ParticipantsNumber" : vm.participantsNumber,
            "PlannedCost":(vm.participantsNumber * 100),
            "MeetingTime": vm.meetingTime.toString(),
            "MeetingRoomName" : vm.selectedmeetingRoomName,
            "MeetingDate": vm.meetingDate,
            "FoodCategory": vm.selectedfoodCategory,
            "FoodSetMenu" : vm.selectedfoodSetMenu,
            "CostCenter": vm.costCenter.toString(),
            "GLNumber" : vm.gLNumber.toString(),
            "PendingWithId": ApprovalChain.LineManager, // 1st approver-
            "Status": ApprovalStatus.Submitted,
            "OnBehalfId": OnBehalf,
        };
        console.log("Chcek", JSON.stringify(Food_data));
        vm.loading = true;
        $('#SuccessModal').modal('show');



        const addPendingApproval = async (FoodId) => {
            try {
                const PendingApproval_data = {
                    "__metadata": { "type": "SP.Data.PendingApprovalListItem" },
                    "Title": `F-${FoodId}`,
                    "PendingWithId": ApprovalChain.LineManager, // 1st approver
                    "RequestLink": `${ABS_URL}/SitePages/food.aspx?FoodId=${uuid.v4()}`,
                    "RequestedById": userId,
                    "ProcessName": "Food",
                    "Status": ApprovalStatus.Submitted,
                };
                await AddListItem("PendingApproval", PendingApproval_data);
            } catch (error) {
                console.error('Error adding PendingApproval:', error);
            }
        };

        const addFoodRequest = async () => {
            try {
                const res = await AddListItem("Food", Food_data);
                await addPendingApproval(res.Id);
                await addHistory(res.Id, ApprovalStatus.Submitted);
                await SaveAllAttachments(res.Id);

                $scope.$apply(() => {
                    vm.loading = false;
                    vm.showMCloseBtn = false;
                    vm.ModalMessage = `Your request has been submitted successfully. ID #${res.Id}`;
                    console.log('FoodRequest added successfully:', res);
                });
            } catch (error) {
                $scope.$apply(() => {
                    vm.loading = false;
                    vm.showMCloseBtn = true;
                    vm.ModalMessage = `An error occurred while submitting your request. Please try again later.`;
                    console.error('Error adding FoodRequest:', error);
                });
            }
        }

        addFoodRequest();
    };
    const addHistory = async (FoodId, Status) => {
        try {
            const History_data = {
                "__metadata": { "type": "SP.Data.FoodLogListItem" },
                "Title": `F-${FoodId}`,
                "FoodId": FoodId,
                "Status": Status,
                "Comment": vm.comment
            };
            await AddListItem("FoodLog", History_data);
        } catch (error) {
            console.error('Error adding History:', error);
        }
    };
    const AddAttachment = async (file, RequestId) => {
        const ListName = "FoodAttachment";
        const data = {
            'Title': `F-${RequestId}`,
            'FoodId': RequestId,
            '__metadata': { "type": "SP.Data.FoodAttachmentListItem" },
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
        const updateFoodRequest = async () => {
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

                let Food_data = {
                    "__metadata": { "type": "SP.Data.FoodListItem" },
                    "PendingWithId": NewPendingWithId,
                    "Status": StatusOnAction
                };

                if (CurerntPendingWithId === ApprovalChain.FoodInCharge) {
                   
                    Food_data = {
                        ...Food_data,
                        "GLNumber": vm.gLNumber.toString(),
                        "ActualCost": vm.actualCost,
                    }

                }

                const PendingApproval_data = {
                    "__metadata": { "type": "SP.Data.PendingApprovalListItem" },
                    "PendingWithId": NewPendingWithId,
                    "Status": StatusOnAction,
                };
                await UpdateListItem("Food", FoodId, Food_data);
                await UpdateListItem("PendingApproval", PendingApprovalId, PendingApproval_data);
                await addHistory(FoodId, StatusOnAction);
                await SaveAllAttachments(FoodId);
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
                    console.error('Error adding FoodRequest:', error);
                });
            }
        }
        updateFoodRequest();
    };
    vm.updateRequest = function () {
        const updateFoodRequest = async () => {
            try {
                vm.loading = true;
                $('#SuccessModal').modal('show');
                
                /* Food dm */
                const Food_data = {
                    "__metadata": { "type": "SP.Data.FoodListItem" },
                    "Agenda": vm.agenda,
                    "ParticipantsNumber" : vm.participantsNumber,
                    "PlannedCost": vm.plannedCost,
                    "MeetingTime": vm.meetingTime,
                    "MeetingRoomName" : vm.selectedmeetingRoomName,
                    "MeetingDate": vm.meetingDate,
                    "FoodCategory": vm.selectedfoodCategory,
                    "FoodSetMenu" : vm.selectedfoodSetMenu,
                    "CostCenter": vm.costCenter,
                    "GLNumber" : vm.gLNumber,

                    "PendingWithId": ApprovalChain.LineManager, // 1st approver-
                    "Status": ApprovalStatus.Submitted,
                    "OnBehalfId": OnBehalf,
                };
                await UpdateListItem("Food", FoodId, Food_data);

                /* PendingApproval dm */
                const PendingApproval_data = {
                    "__metadata": { "type": "SP.Data.PendingApprovalListItem" },
                    "PendingWithId": NewPendingWithId,
                    "Status": ApprovalStatus.Updated,
                };
                await UpdateListItem("PendingApproval", PendingApprovalId, PendingApproval_data);

                await addHistory(FoodId, ApprovalStatus.Updated);
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
                    console.error('Error adding FoodRequest:', error);
                });
            }
        }
        updateFoodRequest();
    }
});
