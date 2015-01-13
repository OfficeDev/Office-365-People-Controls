var initialized = false;
function InitPeoplePicker(element, options) {
	if (options === void 0) { options = {}; }
	if (!initialized) {
		var access = {};
		window.Access = access;
		access.ControlTelemetryAdapter = function () {
			this.writeDiagnosticLog = function () {
			};
			this.writeCustomerActionLog = function () {
			};
			this.startCustomerActionLog = function () {
			};
			this.endCustomerActionLog = function () {
			};
			this.writeInformationalLog = function () {
			};
			this.getCorrelationKey = function () {
				return '';
			};
			this.startPerformance = function () {
			};
			this.endPerformance = function () {
			};
			this.uploadStoreToServer = function () {
			};
			this.logSingletonCustomerAction = function () {
			};
		};
		access.ControlTelemetryAdapter.getStaticTelemetryAdapter = function () { return new access.ControlTelemetryAdapter(); };
		access.TelemetryManager = {};
		access.TelemetryManager.generateGuid = function () {
			return '6efb8c9c-e92b-4588-8273-1f4d7f28267e';
		};
		access.TelemetryManager.get_contextManager = function () {
			return {
				storeCrossScopeCorrelationId: function () {
				},
				getCrossScopeCorrelationId: function () {
				},
			};
		};
		window.SP = {};
		Office.Controls.Runtime.initialize({ sharePointHostUrl: "", appWebUrl: "" });
		initialized = true;
	}
}