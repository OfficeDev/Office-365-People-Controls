/// <reference path="Office.Controls.PeoplePicker.d.ts" />
var params = {
    allowMultipleSelections: true,
    startSearchCharLength: 1,
    delaySearchInterval: 300,
    enableCache: true,
    numberOfResults: 30,
    inputHint: 'Type keyword to search...',
    showValidationErrors: true,
    onAdd: onPersonAddOrRemove,
    onRemove: onPersonAddOrRemove,
    onChange: undefined,
    onFocus: undefined,
    onBlur: undefined,
    onError: undefined,
    resourceStrings: {
        PeoplePickerNoMatch: '没有匹配项',
        PeoplePickerShowingTopNumberOfResults: '显示前{0}条结果',
        PeoplePickerServerProblem: '不能连接到服务器',
        PeoplePickerDefaultMessagePlural: '键入名字或邮件地址...',
        PeoplePickerMultipleMatch: '匹配到多个结果，点击选择',
        PeoplePickerNoResult: '没有匹配',
        PeoplePickerSingleResult: '显示{0}条结果',
        PeoplePickerMultipleResults: '显示{0}条结果',
        PeoplePickerSearching: '正在搜索...',
        PeoplePickerRemovePerson: '移除人或组{0}',
        PeoplePickerDefaultMessage: '键入名字或邮件地址...',
        PeoplePickerSearchResultRecentGroup: '最近',
        PeoplePickerSearchResultMoreGroup: '更多'
    }
};

function onPersonAddOrRemove(control) {
    document.getElementById('ppc_multiple_error').innerHTML = "";
    var people = 'Added people: ';
    control.getAddedPeople().forEach(function (e) {
        people += '<p>{' + e.DisplayName + ', id=' + e.PersonId + '}</p>';
    });
    document.getElementById('ppc_multiple_people').innerHTML = "<pre>" + people + "</pre>";
}
//# sourceMappingURL=sample.js.map