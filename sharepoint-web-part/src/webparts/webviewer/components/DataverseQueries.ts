import { AadHttpClient, IHttpClientOptions } from "@microsoft/sp-http";

import { ISignatureRequestProps } from "./ISignatureRequestProps";
import { IUserAccountsProps } from "./IUserAccountsProps";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export default class DataverseQueries implements ISignatureRequestProps, IUserAccountsProps {

	public currentUser: IUserAccountsProps;
	public users: IUserAccountsProps[];
	public context: WebPartContext;

	public accessToken: string;

	private _requestHeaders: Headers = new Headers();

	public constructor(props: WebPartContext) {
		this.context = props;
		this._requestHeaders.append('Content-Type', 'application/json');
	}

	public async Login(): Promise<void> {
		// console.log("Using Login() method");
		const options: IHttpClientOptions = {
			headers: this._requestHeaders
		}

		// Using Graph here, but any 1st or 3rd party REST API that requires Azure AD auth can be used here.
		await this.context.aadHttpClientFactory
			.getClient(process.env.CRM_URL)
			.then((client: AadHttpClient) => {
				return client
					.get(
						`${process.env.CRM_URL}/api/data/v9.2/cr80b_signaturerequests(ceed5dc6-357a-ed11-81ad-00224826545c)?$select=cr80b_requestor`,
						AadHttpClient.configurations.v1,
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
}