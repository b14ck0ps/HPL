/// <reference path="../../JSDependency/BaseModule.js" />

const ItAssetApp = angular.module('ItAssetApp', [])
const AssetId = new URLSearchParams(window.location.search).get('AssetId');
let LastAssetNumber;

ItAssetApp.controller('AssetFormController', function ($scope, $timeout) {
    const vm = this;
    vm.loading = true;
    vm.showUpdateButton = false;
    vm.formData = {};
    const fetchData = async () => {
        vm.subCategories = [{ name: 'Laptops' }, { name: 'Desktops' }], vm.acquisitionTypes = [{ name: 'As Own' }, { name: 'Support' }, { name: 'General' }, { name: 'Sold' }, { name: 'Damage' }, { name: 'Support Store' }, { name: 'Lost' }, { name: 'Disposed' }], vm.categories = [{ name: 'Computer' }], vm.brands = [{ name: 'Dell' }, { name: 'HP' }, { name: 'Canon' }], vm.operations = [{ name: 'MIS' }, { name: 'SMD' }, { name: 'Accounts' }, { name: 'HRD' }], vm.assetLocations = [{ name: 'HPL HQ' }, { name: 'Khulna' }, { name: 'Sylhet' }, { name: 'Rajshahi' }];

        $scope.$watchGroup(['vm.formData.subCategory', 'vm.formData.purchaseDate', 'vm.formData.department'], function () {
            let _year = vm.formData.purchaseDate?.getFullYear();
            let _categoryPrefix = (vm.formData.subCategory === '' || vm.formData.subCategory === undefined) ? '' : (vm.formData.subCategory.includes('Laptop') ? 'LAP' : 'DES');
            /* let _departmentPrefix = (vm.formData.department || '').split(' ').map(word => word.charAt(0)).join(''); */
            let _departmentPrefix = vm.formData.department ?? '';
            const _assetNumber = LastAssetNumber === undefined ? '' : String(LastAssetNumber + 1).padStart(3, '0');

            const tagParts = ['HPL', 'IE', _categoryPrefix, _year, _departmentPrefix, _assetNumber].filter(Boolean);
            if (!AssetId)
                vm.formData.tag = tagParts.slice(0, 5).join('-') + tagParts.slice(5).join('');
        });

        try {
            const LastAssetRow = await GetByList('ItAssetMaster', '$select=Id&$orderby=Id desc&$top=1');
            LastAssetNumber = LastAssetRow[0].ID;
            const userdata = await GetByList("Employees", `$select=Code,name,deptId,phone,designation&$filter=Code eq ${userId}`);
            const fetchAssetUser = async () => {
                try {
                    const query = `$select=Code,name,deptId,phone,designation,email/Id,email/Title&$expand=email,email/Id`;
                    const data = await GetByList('Employees', query);
                    $scope.$apply(() => {
                        vm.AssetUsers = data.map(item => {
                            return {
                                name: item.email.Title,
                                email: item.email.Id,
                            };
                        });
                        vm.SelectedAssetUser = "";
                    });
                } catch (error) {
                    console.error(error);
                }
            };
            const fetchDepartment = async () => {
                try {
                    const query = `$select=Name,DepartmentTag`;
                    const data = await GetByList('Department', query);
                    $scope.$apply(() => {
                        vm.departments = data.map(item => {
                            return {
                                name: item.Name,
                                tag: item.DepartmentTag
                            };
                        });
                        vm.department = "";
                    });
                } catch (error) {
                    console.error(error);
                }
            };
            await fetchAssetUser();
            await fetchDepartment();

            if (AssetId) {
                fetchAssetData(AssetId);
                vm.showUpdateButton = true;
            }

            $scope.$apply(() => {
                vm.userInfo = userdata[0];
                vm.loading = false;
            });
        } catch (error) {
            console.error(error);
        }
    };

    fetchData();

    vm.saveAsset = async () => {

        if (vm.formData.SelectedAssetUser === undefined || vm.formData.SelectedAssetUser === '') {
            alert('Please select an asset user.');
            return;
        }

        try {
            const data = {
                __metadata: { type: 'SP.Data.ItAssetMasterListItem' },
                Title: 'IT Asset',
                ...vm.formData,
                AssetUsersId: vm.formData.SelectedAssetUser
            };

            const { SelectedAssetUser, ...filteredData } = data;
            vm.loading = true;
            $('#SuccessModal').modal('show');
            const result = await AddListItem('ItAssetMaster', filteredData);

            if (result) {
                $scope.$apply(() => {
                    vm.loading = false;
                    vm.showMCloseBtn = false;
                    vm.ModalMessage = `Your request has been submitted successfully. Tag #${vm.formData.tag}`;
                });
            }
        } catch (error) {
            $scope.$apply(() => {
                vm.loading = false;
                vm.showMCloseBtn = true;
                vm.ModalMessage = `An error occurred while submitting your request. Please try again later.`;
                console.error(error);
            });
        }
    };

    const fetchAssetData = async (id) => {
        try {
            const query = `$select=*&$filter=ID eq ${id}`;
            const response = await GetByList('ItAssetMaster', query, id);
            const data = response[0];

            $scope.$apply(() => {
                vm.formData = {
                    ...data,
                    purchaseDate: new Date(data.purchaseDate),
                    currentUserRequisitionDate: new Date(data.currentUserRequisitionDate),
                    warrantyPeriodFrom: new Date(data.warrantyPeriodFrom),
                    warrantyPeriodTo: new Date(data.warrantyPeriodTo)
                }
                vm.showUpdateButton = true;
                $timeout(function () {
                    $("#AssetUserName").val(data.AssetUsersId).trigger("change");
                });
                vm.loading = false;
            });

        } catch (error) {
            console.error(error);
        }
    };

    vm.updateAsset = async () => {
        const selectedProperties = ['assetNumber', 'subCategory', 'acquisitionType', 'purchaseVendor', 'category', 'brand', 'model', 'number', 'assetTitle', 'purchaseDate', 'manufacturer', 'purchasePrice', 'tag', 'name', 'version', 'securityPatch', 'usefulLife', 'currentUserRequisitionDate', 'warrantyPeriodFrom', 'productSLNo', 'warrantyPeriodTo', 'position', 'email', 'employeeId', 'operation', 'workOrderNumber', 'assetLocation', 'assetOwner', 'department', 'assetCustodian'];
        const filteredFormData = {};
        selectedProperties.forEach(property => {
            if (vm.formData[property] !== undefined) {
                filteredFormData[property] = vm.formData[property];
            }
        });

        vm.loading = true;
        $('#SuccessModal').modal('show');
        try {
            await UpdateListItem('ItAssetMaster', AssetId, { __metadata: { type: 'SP.Data.ItAssetMasterListItem' }, ...filteredFormData });
            vm.loading = false;
            $scope.$apply(() => {
                vm.showMCloseBtn = false;
                vm.ModalMessage = `Asset ${vm.formData.tag} has been Updated successfully.`;
                vm.loading = false;
            });

        } catch (error) {
            $scope.$apply(() => {
                vm.showMCloseBtn = true;
                vm.ModalMessage = `An error occurred while updating. Please try again later.`;
                vm.loading = false;
            });
        }
    }
    vm.RedirectToDashboard = () => window.location.href = `${ABS_URL}/SitePages/ItAssetDashboard.aspx`;

    vm.testMe = () => { const getRandomString = (length) => { const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'; let result = ''; for (let i = 0; i < length; i++) { result += characters.charAt(Math.floor(Math.random() * characters.length)) } return result }, getRandomDate = (start, end) => { return new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime())) }, getRandomArrayElement = (array) => { return array[Math.floor(Math.random() * array.length)] }; vm.formData.assetNumber = getRandomString(8), vm.formData.subCategory = getRandomArrayElement(vm.subCategories).name, vm.formData.acquisitionType = getRandomArrayElement(vm.acquisitionTypes).name, vm.formData.purchaseVendor = getRandomString(15), vm.formData.category = getRandomArrayElement(vm.categories).name, vm.formData.brand = getRandomArrayElement(vm.brands).name, vm.formData.model = getRandomString(10), vm.formData.number = getRandomString(8), vm.formData.assetTitle = getRandomString(15), vm.formData.purchaseDate = getRandomDate(new Date(2022, 0, 1), new Date()), vm.formData.manufacturer = getRandomString(10), vm.formData.purchasePrice = Math.floor(Math.random() * 1000) + 1, vm.formData.name = getRandomString(10), vm.formData.version = getRandomString(5), vm.formData.securityPatch = getRandomString(5), vm.formData.usefulLife = Math.floor(Math.random() * 10) + 1, vm.formData.currentUserRequisitionDate = getRandomDate(new Date(2022, 0, 1), new Date()), vm.formData.warrantyPeriodFrom = getRandomDate(new Date(2022, 0, 1), new Date()), vm.formData.productSLNo = getRandomString(8), vm.formData.warrantyPeriodTo = getRandomDate(new Date(2023, 0, 1), new Date()), vm.formData.position = getRandomString(10), vm.formData.email = getRandomString(10) + "@example.com", vm.formData.employeeId = getRandomString(8), vm.formData.operation = getRandomArrayElement(vm.operations).name, vm.formData.workOrderNumber = getRandomString(8), vm.formData.assetLocation = getRandomArrayElement(vm.assetLocations).name, vm.formData.assetOwner = getRandomString(10), vm.formData.department = getRandomArrayElement(vm.departments).tag, vm.formData.assetCustodian = getRandomString(10); console.log('Random data generated for testing:', vm.formData) };
});