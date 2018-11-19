import {
    SPHttpClient
} from '@microsoft/sp-http';
export default interface IPrintDialogContentProps {
    close: () => void;
    httpClient: SPHttpClient;
    webUrl: string;
}