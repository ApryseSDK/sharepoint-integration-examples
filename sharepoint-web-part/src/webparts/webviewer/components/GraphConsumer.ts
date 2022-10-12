import { AadHttpClient } from "@microsoft/sp-http";

import { IUserData } from "./IUserData";
import { IGraphConsumerProps } from "./IGraphConsumerProps";
import { IGraphConsumerState } from "./IGraphConsumerState";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export default class GraphConsumer implements IGraphConsumerProps, IGraphConsumerState {

	public currentUser: IUserData;
	public users: IUserData[];
	public context: WebPartContext;

	public constructor(props: WebPartContext) {
		this.context = props;
	}

	public async GetCurrentUser(): Promise<void> {
		console.log("Using GetCurrentUser() method");

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
				console.log(json);

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
		console.log("Using ListUsers() method");

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
				console.log(json);

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

}