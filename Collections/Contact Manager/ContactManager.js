/* contact manager */
jQuery(document).ready(function(){
	jQuery("body").on( "focus", ".cmTextInclude", function() {
		jQuery("#" + this.name + "_r").prop("checked", true);
	});
})