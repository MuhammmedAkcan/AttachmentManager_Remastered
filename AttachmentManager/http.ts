export const http = async (request: Request | string): Promise<any> => {
    return new Promise(resolve => {
        fetch(request)
            .then(response => response.json())
            .then(body => {
                resolve(body);
            });
    });
};