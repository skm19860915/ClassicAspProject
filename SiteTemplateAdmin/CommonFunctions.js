function copySelectToHidden(selectFieldID, hiddenFieldID, copyAll) {
	var selectField = document.getElementById(selectFieldID);
	var hiddenField = document.getElementById(hiddenFieldID);

	var hiddenFieldValue = '';

	for(var i = 0; i < selectField.options.length; i++) {
		if(copyAll || selectField.options[i].selected) {
			if(hiddenFieldValue != '')
					hiddenFieldValue += ',';

			hiddenFieldValue += selectField.options[i].value;
		}
	}

	hiddenField.value = hiddenFieldValue;
}

function copyHiddenToSelect(hiddenFieldID, selectFieldID) {
	var hiddenField = document.getElementById(hiddenFieldID);
	var selectField = document.getElementById(selectFieldID);

	var hiddenFieldValue = ',' + hiddenField.value + ',';

	for(var i = 0; i < selectField.options.length; i++)
		if(hiddenFieldValue.indexOf(',' + selectField.options[i].value + ',') >= 0)
			 selectField.options[i].selected = true;
}

function moveOptions(fromSelectBoxID, toSelectBoxID) {
        fromSelectBox = document.getElementById(fromSelectBoxID);
        toSelectBox = document.getElementById(toSelectBoxID);
        
        for(var i = fromSelectBox.options.length - 1; i >= 0; i--)
                if(fromSelectBox.options[i].selected)
                        toSelectBox.insertBefore(fromSelectBox.options[i], null);
}

function toggleFormFieldDisplay(formFieldID, showFormField, resetFieldValue, displayFieldID) {
	var formField = document.getElementById(formFieldID);
	var displayField = document.getElementById(displayFieldID);
	if(!displayField)
		displayField = formField;
	
	if(showFormField) {
		displayField.style.display = '';
	} else {
		displayField.style.display = 'none';
		if(formField.options)
			selectOption(formField, resetFieldValue);
		else
			formField.value = resetFieldValue;
	}
}

function selectOption(selectBox, optionValueList) {
	optionValueList = ',' + optionValueList + ',';

	for(var i = 0; i < selectBox.options.length; i++) {
		if(optionValueList.indexOf(',' + selectBox.options[i].value + ',') >= 0) {
			selectBox.options[i].selected = true;
			if(selectBox.type == 'select-one')
				return;
		}
	}
}