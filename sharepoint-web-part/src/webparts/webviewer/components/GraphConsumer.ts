import { AadHttpClient, IHttpClientOptions } from "@microsoft/sp-http";

import { IUserData } from "./IUserData";
import { IGraphConsumerProps } from "./IGraphConsumerProps";
import { IGraphConsumerState } from "./IGraphConsumerState";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export default class GraphConsumer implements IGraphConsumerProps, IGraphConsumerState {

	public currentUser: IUserData;
	public users: IUserData[];
	public context: WebPartContext;
	public testFile: Blob;
	public xfdf: string;
	private _requestHeaders: Headers = new Headers();


	public constructor(props: WebPartContext) {
		this.context = props;
		this._requestHeaders.append('OData-MaxVersion', '4.0');
		this._requestHeaders.append('OData-Version', '4.0');
		this._requestHeaders.append('Content-Type', 'application/json; charset=utf-8');
		this._requestHeaders.append('Accept', 'application/json');
		this._requestHeaders.append('Prefer', 'odata.include-annotations=*');
	}

	public async GetCurrentUser(): Promise<void> {
		// console.log("Using GetCurrentUser() method");

		// Using Graph here, but any 1st or 3rd party REST API that requires Azure AD auth can be used here.
		await this.context.aadHttpClientFactory
			.getClient("https://graph.microsoft.com")
			.then((client: AadHttpClient) => {
				return client
					.get(
						`https://graph.microsoft.com/v1.0/me`,
						AadHttpClient.configurations.v1
					);
			})
			.then(response => {
				return response.json();
			})
			.then(json => {
				// Log the result in the console for testing purposes
				// console.log(json);

				const userData: IUserData = ({
					displayName: json.displayName,
					mail: json.mail,
					userPrincipalName: json.userPrincipalName,
				});

				this.currentUser = userData;
			})
			.catch(error => {
				console.error(error);
			});
	}

	public async ListUsers(): Promise<void> {
		// console.log("Using ListUsers() method");

		// Using Graph here, but any 1st or 3rd party REST API that requires Azure AD auth can be used here.
		await this.context.aadHttpClientFactory
			.getClient("https://graph.microsoft.com")
			.then((client: AadHttpClient) => {
				return client
					.get(
						`https://graph.microsoft.com/v1.0/users`,
						AadHttpClient.configurations.v1
					);
			})
			.then(response => {
				return response.json();
			})
			.then(json => {

				// Prepare the output array
				const usersArr: Array<IUserData> = new Array<IUserData>();

				// Log the result in the console for testing purposes
				// console.log(json);

				// Map the JSON response to the output array
				json.value.map((item: IUserData) => {
					usersArr.push({
						displayName: item.displayName,
						mail: item.mail,
						userPrincipalName: item.userPrincipalName,
					});
				});

				// Update the component state accordingly to the result
				this.users = usersArr;
			})
			.catch(error => {
				console.error(error);
			});
	}

	public async TestAPI(UniqueId: string): Promise<void> {
		// console.log("Using TestAPI() method");

		// Using Graph here, but any 1st or 3rd party REST API that requires Azure AD auth can be used here.
		await this.context.aadHttpClientFactory
			.getClient(process.env.CRM_URL)
			.then((client: AadHttpClient) => {
				return client
					.get(
						`${process.env.CRM_URL}/api/data/v9.2/cr80b_signaturerequests(${UniqueId})`,
						AadHttpClient.configurations.v1
					);
			})
			.then(response => {
				return response.json();
			})
			.then(json => {
				// Log the result in the console for testing purposes
				// console.log(json);

				// Map the JSON response to the output array
				this.xfdf = json.cr80b_xfdf;
			})
			.catch(error => {
				console.error(error);
			});
	}

	public async SaveXFDF(XFDF: string, UniqueId: string): Promise<void> {
		// console.log("Using SaveXFDF() method headers:" + this._requestHeaders);
		// eslint-disable-next-line @typescript-eslint/typedef
		let record = { cr80b_xfdf: ""};
		record.cr80b_xfdf = XFDF; // Text

		const options: IHttpClientOptions = {
			method: "PATCH",
			headers: this._requestHeaders,
			body: JSON.stringify(record)
		}

		// Using Graph here, but any 1st or 3rd party REST API that requires Azure AD auth can be used here.
		await this.context.aadHttpClientFactory
			.getClient(process.env.CRM_URL)
			.then((client: AadHttpClient) => {
				return client
					.fetch(
						`${process.env.CRM_URL}/api/data/v9.2/cr80b_signaturerequests(${UniqueId})`,
						AadHttpClient.configurations.v1,
						options
					);
			})
			.then(response => {
				return response.json();
			})
			.then(json => {
				// Log the result in the console for testing purposes
				// console.log(json);
				
			})
			.catch(error => {
				console.error(error);
			});
	}

	public async GetFileAPI(UniqueId: string): Promise<void> {
		// console.log("Using GetFileAPI() method");

		// Using Graph here, but any 1st or 3rd party REST API that requires Azure AD auth can be used here.
		await this.context.aadHttpClientFactory
			.getClient(process.env.CRM_URL)
			.then((client: AadHttpClient) => {
				return client
					.get(
						`${process.env.CRM_URL}/api/data/v9.2/cr80b_signaturerequests(${UniqueId})/cr80b_filetosign/$value`,
						AadHttpClient.configurations.v1
					);
			})
			.then(response => {
				return response.blob();
			})
			.then(blob => {
				// Log the result in the console for testing purposes
				// console.log(blob);
				this.testFile = blob;
				
			})
			.catch(error => {
				console.error(error);
			});
	}
}