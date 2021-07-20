var wb;
function importf(obj) {
	if(!obj.files) {
		return;
	}
	var f = obj.files[0];
	var reader = new FileReader();
	reader.onload = function(e) {
		var data = e.target.result;
		wb = XLSX.read(data, {
			type: 'binary'
		});
		var data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:""});
		$("#tableWithDataFromExcel thead tr").append("<th id='CheckboxesHeaderId' style='position: sticky; right: 0;top:0;background:#1C6EA4;border-bottom: 2px solid #444444; z-index: 3'>Select to create</th>");
		
	};
	reader.readAsBinaryString(f)
	var checkExist = setInterval(function() {
		if ($("#tableWithDataFromExcel tbody tr td").length) {
			
			$("#tableWithDataFromExcel tbody tr").each(function(){
				$(this).find("td").each(function(){
					var curval = $(this).html().toString();
					$(this).html("<input type='text' value='"+curval+"'>");
				});
				$(this).append('<td style="right: 0; background: #b1d4f2; position: sticky; z-index: 0"><input type="checkbox" checked ></td>');
			});
			$('#CheckboxesHeaderId').click(function()
			{
				$("#tableWithDataFromExcel tbody tr td:last-child input[type='checkbox']").each(function(){
					$(this).click();
				})
			});
			$('#tableWithDataFromExcel').addClass("paleBlueRows");
			$("#messageResponse").css("visibility","hidden");
		   clearInterval(checkExist);
		}
	 }, 100)
	
	$("#buttonCreateRecocrdsId").css("visibility","visible");
	$("#loadDropdownsButtonId").css("visibility","visible");
	$("#changeAttributesButtonId").css("visibility","visible");
}