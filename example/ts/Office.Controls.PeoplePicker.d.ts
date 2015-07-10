declare module Office {
    export module Controls {
        export class PeoplePickerRecord {
            DisplayName: string;
            Description: string;
            PersonId: string;
            ImgSrc: string;
            constructor();
        }
        export class ValidationError {
            errorName: string;
            localizedErrorMessage: string;
        }
        export interface DataProvider {
            getPrincipals(keyword: string, callback: (error: string, results: Office.Controls.PeoplePickerRecord[]) => void): void;
            getImageAsync(personId: string, callback: (error: string, imgSrc: string) => void): void;
        }
        export interface PeoplePickerOptions {
            allowMultipleSelections?: boolean;
            startSearchCharLength?: number;
            delaySearchInterval?: number;
            enableCache?: boolean;
            numberOfResults?: number;
            inputHint?: string;
            showValidationErrors?: boolean;
            showImage?: boolean;
            onAdd?: (control: Office.Controls.PeoplePicker, person: Office.Controls.PeoplePickerRecord) => void;
            onRemove?: (control: Office.Controls.PeoplePicker, person: Office.Controls.PeoplePickerRecord) => void;
            onChange?: (control: Office.Controls.PeoplePicker) => void;
            onFocus?: (control: Office.Controls.PeoplePicker) => void;
            onBlur?: (control: Office.Controls.PeoplePicker) => void;
            onError?: (control: Office.Controls.PeoplePicker, error: ValidationError) => void;
            resourceStrings?: any;
        }
        export class PeoplePicker {
            allowMultipleSelections: boolean;
            startSearchCharLength: number;
            delaySearchInterval: number;
            enableCache: boolean;
            numberOfResults: number;
            inputHint: string;
            showValidationErrors: boolean;
            showImage?: boolean;
            onAdd: (control: Office.Controls.PeoplePicker, person: Office.Controls.PeoplePickerRecord) => void;
            onRemove: (control: Office.Controls.PeoplePicker, person: Office.Controls.PeoplePickerRecord) => void;
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
            constructor(root: HTMLElement, dataProvider: Office.Controls.DataProvider, options: PeoplePickerOptions);
            static create(root: HTMLElement, dataProvider: Office.Controls.DataProvider, options: PeoplePickerOptions): Office.Controls.PeoplePicker;
        }
    }
}