declare module Office {
    export module Controls {
        export class PeoplePickerRecord {
            DisplayName: string;
            Description: string;
            PersonId: string;
            constructor();
        }
        export class ValidationError {
            errorName: string;
            localizedErrorMessage: string;
        }
        export interface DataProvider {
            getPrincipals(keyword: string, callback: (error: string, results: Office.Controls.PeoplePickerRecord[]) => void): void;
        }
        export interface PeoplePickerOptions {
            allowMultipleSelections?: boolean;
            startSearchCharLength?: number;
            delaySearchInterval?: number;
            enableCache?: boolean;
            numberOfResults?: number;
            inputHint?: string;
            showValidationErrors?: boolean;
            onAdded?: (control: Office.Controls.PeoplePicker, person: Office.Controls.PeoplePickerRecord) => void;
            onRemoved?: (control: Office.Controls.PeoplePicker, person: Office.Controls.PeoplePickerRecord) => void;
            onChange?: (control: Office.Controls.PeoplePicker) => void;
            onFocus?: (control: Office.Controls.PeoplePicker) => void;
            onBlur?: (control: Office.Controls.PeoplePicker) => void;
            onError?: (control: Office.Controls.PeoplePicker, error: ValidationError) => void;
            resourceStrings?: any;
        }
        export class PeoplePicker {
            allowMultiple: boolean;
            startSearchCharLength: number;
            delaySearchInterval: number;
            enableCache: boolean;
            numberOfResults: number;
            inputHint: string;
            showValidationErrors: boolean;
            onAdded: (control: Office.Controls.PeoplePicker, person: Office.Controls.PeoplePickerRecord) => void;
            onRemoved: (control: Office.Controls.PeoplePicker, person: Office.Controls.PeoplePickerRecord) => void;
            onChange: (control: Office.Controls.PeoplePicker) => void;
            onFocus: (control: Office.Controls.PeoplePicker) => void;
            onBlur: (control: Office.Controls.PeoplePicker) => void;
            onError: (control: Office.Controls.PeoplePicker, error: ValidationError) => void;
            dataProvider: Office.Controls.DataProvider;
            showInputHint: boolean;
            getAddedPeople: Office.Controls.PeoplePickerRecord[];
            hasErrors: boolean;
            reset(): void;
            remove(entryToRemove: Office.Controls.PeoplePickerRecord): void;
            add(input: string): void;
            add(info: Office.Controls.PeoplePickerRecord): void;
            add(info: Office.Controls.PeoplePickerRecord, resolve: boolean): void;
            clearCacheData(): void;
            getErrorDisplayed(): Office.Controls.ValidationError;
            constructor(root: HTMLElement, dataProvider: Office.Controls.DataProvider);
            constructor(root: HTMLElement, dataProvider: Office.Controls.DataProvider, parameterObject: PeoplePickerOptions);
            static create(root: HTMLElement, dataProvider: Office.Controls.DataProvider, parameterObject: PeoplePickerOptions): Office.Controls.PeoplePicker;
        }
    }
}