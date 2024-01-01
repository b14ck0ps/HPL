/// <reference path="../../JSDependency/BaseModule.js" />

const ItAssetApp = angular.module('ItAssetApp', [])
let LastAssetNumber;

ItAssetApp.controller('AssetFormController', function ($scope) {
    const vm = this;
    vm.loading = true;
    vm.showUpdateButton = false;

    const fetchUserData = async () => {
        vm.subCategories = [
            { name: 'Laptops' },
            { name: 'Desktops' }
        ];

        vm.acquisitionTypes = [
            { name: 'As Own' },
            { name: 'Support' },
            { name: 'General' },
            { name: 'Sold' },
            { name: 'Damage' },
            { name: 'Disposed' }
        ];

        vm.categories = [
            { name: 'Computer' }
        ];

        vm.brands = [
            { name: 'Dell' },
            { name: 'HP' },
            { name: 'Canon' }
        ];

        vm.operations = [
            { name: 'MIS' },
            { name: 'SMD' },
            { name: 'Accounts' },
            { name: 'HRD' }
        ];

        vm.assetLocations = [
            { name: 'HPL HQ' },
            { name: 'Khulna' },
            { name: 'Sylhet' },
            { name: 'Rajshahi' }
        ];

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
            fetchAssetUser();
            fetchDepartment();
            $scope.$apply(() => {
                vm.userInfo = userdata[0];
                vm.loading = false;
            });
        } catch (error) {
            console.error(error);
        }
    };

    fetchUserData();
    $scope.$watchGroup(['vm.subCategory', 'vm.purchaseDate', 'vm.department'], function () {
        let _year = vm.purchaseDate?.getFullYear();
        let _categoryPrefix = (vm.subCategory === '' || vm.subCategory === undefined) ? '' : (vm.subCategory.includes('Laptop') ? 'LAP' : 'DES');
        /* let _departmentPrefix = (vm.department || '').split(' ').map(word => word.charAt(0)).join(''); */
        let _departmentPrefix = vm.department ?? '';
        const _assetNumber = LastAssetNumber === undefined ? '' : String(LastAssetNumber + 1).padStart(3, '0');

        const tagParts = ['HPL', 'IE', _categoryPrefix, _year, _departmentPrefix, _assetNumber].filter(Boolean);
        vm.tag = tagParts.slice(0, 5).join('-') + tagParts.slice(5).join('');
    });

    vm.saveAsset = async () => {

        if (vm.SelectedAssetUser === undefined || vm.SelectedAssetUser === '') {
            alert('Please select an asset user.');
            return;
        }

        try {
            const data = {
                __metadata: { type: 'SP.Data.ItAssetMasterListItem' },
                Title: 'IT Asset',
                assetNumber: vm.assetNumber,
                subCategory: vm.subCategory,
                acquisitionType: vm.acquisitionType,
                purchaseVendor: vm.purchaseVendor,
                category: vm.category,
                brand: vm.brand,
                model: vm.model,
                number: vm.number,
                assetTitle: vm.assetTitle,
                purchaseDate: vm.purchaseDate,
                manufacturer: vm.manufacturer,
                purchasePrice: vm.purchasePrice,
                tag: vm.tag,
                name: vm.name,
                version: vm.version,
                securityPatch: vm.securityPatch,
                usefulLife: vm.usefulLife,
                currentUserRequisitionDate: vm.currentUserRequisitionDate,
                warrantyPeriodFrom: vm.warrantyPeriodFrom,
                productSLNo: vm.productSLNo,
                warrantyPeriodTo: vm.warrantyPeriodTo,
                AssetUsersId: vm.SelectedAssetUser,
                position: vm.position,
                email: vm.email,
                employeeId: vm.employeeId,
                operation: vm.operation,
                workOrderNumber: vm.workOrderNumber,
                assetLocation: vm.assetLocation,
                assetOwner: vm.assetOwner,
                department: vm.department,
                assetCustodian: vm.assetCustodian
            };
            vm.loading = true;
            $('#SuccessModal').modal('show');
            const result = await AddListItem('ItAssetMaster', data);

            if (result) {
                $scope.$apply(() => {
                    vm.loading = false;
                    vm.showMCloseBtn = false;
                    vm.ModalMessage = `Your request has been submitted successfully. Tag #${vm.tag}`;
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

    vm.testMe = () => {
        const getRandomString = (length) => {
            const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
            let result = '';
            for (let i = 0; i < length; i++) {
                result += characters.charAt(Math.floor(Math.random() * characters.length));
            }
            return result;
        };

        const getRandomDate = (start, end) => {
            return new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
        };

        const getRandomArrayElement = (array) => {
            return array[Math.floor(Math.random() * array.length)];
        };

        vm.assetNumber = getRandomString(8);
        vm.subCategory = getRandomArrayElement(vm.subCategories).name;
        vm.acquisitionType = getRandomArrayElement(vm.acquisitionTypes).name;
        vm.purchaseVendor = getRandomString(15);
        vm.category = getRandomArrayElement(vm.categories).name;
        vm.brand = getRandomArrayElement(vm.brands).name;
        vm.model = getRandomString(10);
        vm.number = getRandomString(8);
        vm.assetTitle = getRandomString(15);
        vm.purchaseDate = getRandomDate(new Date(2022, 0, 1), new Date());
        vm.manufacturer = getRandomString(10);
        vm.purchasePrice = Math.floor(Math.random() * 1000) + 1;
        vm.tag = getRandomString(8);
        vm.name = getRandomString(10);
        vm.version = getRandomString(5);
        vm.securityPatch = getRandomString(5);
        vm.usefulLife = Math.floor(Math.random() * 10) + 1;
        vm.currentUserRequisitionDate = getRandomDate(new Date(2022, 0, 1), new Date());
        vm.warrantyPeriodFrom = getRandomDate(new Date(2022, 0, 1), new Date());
        vm.productSLNo = getRandomString(8);
        vm.warrantyPeriodTo = getRandomDate(new Date(2023, 0, 1), new Date());
        vm.position = getRandomString(10);
        vm.email = getRandomString(10) + "@example.com";
        vm.employeeId = getRandomString(8);
        vm.operation = getRandomArrayElement(vm.operations).name;
        vm.workOrderNumber = getRandomString(8);
        vm.assetLocation = getRandomArrayElement(vm.assetLocations).name;
        vm.assetOwner = getRandomString(10);
        vm.department = getRandomArrayElement(vm.departments).tag;
        vm.assetCustodian = getRandomString(10);

        console.log('Random data generated for testing:', vm);
    };


});