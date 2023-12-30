/**
 * The absolute URL of the current SharePoint site.
 * @type {string}
 */
const ABS_URL = _spPageContextInfo.webAbsoluteUrl;
const API_GET_HEADERS = { "Accept": "application/json;odata=verbose" };
const API_POST_HEADERS = {
    "Accept": "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose",
    "X-RequestDigest": $("#__REQUESTDIGEST").val()
};
const API_UPDATE_HEADERS = { ...API_POST_HEADERS, "IF-MATCH": "*", "X-HTTP-Method": "MERGE" };
const getApiEndpoint = (ListName) => `${ABS_URL}/_api/web/lists/getByTitle('${ListName}')/items`;
/**
 * The ID of the current user.
 * @type {number}
 */
const userId = _spPageContextInfo.userId;

/**
 * Retrieves data from a SharePoint list using the SharePoint REST API.
 * @param {string} ListName - The name of the SharePoint list.
 * @param {string} [query] - The query string to filter the results.
 * @returns {Promise<Array>} - A promise that resolves to an array of list items.
 * @throws {Error} - If an error occurs while fetching the data.
 */
const GetByList = async (ListName, query) => {
    try {
        let endpoint = `${getApiEndpoint(ListName)}`

        if (query)
            endpoint += `?${query}`

        const response = await axios.get(endpoint, {
            headers: API_GET_HEADERS
        });
        return response.data.d.results;
    } catch (error) {
        throw error
    }
};

/**
 * Retrieves data from a SharePoint list using the SharePoint REST API.
 * @param {string} ListName - The name of the SharePoint list.
 * @param {object} data - The data to be added to the list.
 * @returns {Promise<Array>} - A promise that resolves to an array of list items.
 * @throws {Error} - If an error occurs while fetching the data.
 */
const AddListItem = async (ListName, data) => {
    try {
        const response = await axios.post(getApiEndpoint(ListName), data, {
            headers: API_POST_HEADERS
        });
        return response.data.d;
    } catch (error) {
        throw error
    }
};

/**
 * Updates a list item in SharePoint.
 *
 * @param {string} ListName - The name of the SharePoint list.
 * @param {number} Id - The ID of the list item to be updated.
 * @param {object} data - The data to be updated in the list item.
 * @returns {Promise<object>} - A promise that resolves to the updated list item data.
 * @throws {Error} - If an error occurs during the update process.
 */
const UpdateListItem = async (ListName, Id, data) => {
    try {
        const response = await axios.post(`${getApiEndpoint(ListName)}(${Id})`, data, {
            headers: API_UPDATE_HEADERS
        });
        return response.data.d;
    } catch (error) {
        throw error
    }
}

/**
 * Uploads a file to SharePoint.
 * @param {number} itemId - The ID of the item in SharePoint.
 * @param {File} file - The file to upload.
 * @param {string} listName - The name of the SharePoint list.
 * @returns {Promise<void>} - A Promise that resolves when the file is uploaded successfully.
 */
const uploadFileToSharePoint = async (itemId, file, listName) => {
    try {
        const buffer = await getFileBufferAsync(file);
        const queryUrl = `${getApiEndpoint(listName)}(${itemId})/AttachmentFiles/add(FileName='${file.name}')`;
        await axios.post(queryUrl, buffer, {
            headers: API_POST_HEADERS
        });
    } catch (error) {
        console.error('Error uploading file:', error);
    }
}

/**
 * Reads the contents of a file and returns it as an ArrayBuffer.
 *
 * @param {File} file - The file to read.
 * @returns {Promise<ArrayBuffer>} A promise that resolves with the file's contents as an ArrayBuffer.
 */
const getFileBufferAsync = (file) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = (e) => reject(e.target.error);
        reader.readAsArrayBuffer(file);
    });
}