$(function() {
	//下载
	$('#J_download').on('click', function() {
		$.ajax({
			url: "/excelExport",
			type: 'post',
			dataType: 'json',
			success: function(response) {
				if (response.code == 1) {
					window.open(response.download);
				}else{
					alert(response.message);
				}
			}
		});
	});

	
	//上传文件
	var uploader = WebUploader.create({
		auto: true,
		server: '/importExcel',
		pick: '#J_upload',
		resize: false,
		accept: {
	        mimeTypes: '.xlsx,.xls'
	    }
	});

	uploader.on('fileQueued', function(file) {
		console.log('上传中...');
	});

	uploader.on('uploadSuccess', function(file, response) {
		console.log('上传成功');
		console.log(response.listData);
	});

	uploader.on('uploadError', function(file) {
		alert('上传出错');
	});

	setTimeout(function(){
		$('input').hide();
	},20);
});