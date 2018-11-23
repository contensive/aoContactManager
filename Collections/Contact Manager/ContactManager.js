/* contact manager */
jQuery(document).ready(function(){
	jQuery("body").on( "focus", ".cmTextInclude", function() {
		jQuery("#" + this.name + "_r").prop("checked", true);
	});
	jQuery("body").on( "click", "#cmSelectAll", function() {
		// jQuery(".cmSelect").prop("checked", true);
		jQuery(".cmSelect").prop("checked", jQuery("#cmSelectAll").prop("checked"));
	});
})