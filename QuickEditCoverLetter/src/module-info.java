module QuickEditCoverLetter {
	requires javafx.controls;
	requires javafx.fxml;
	requires javafx.graphics;
	requires org.apache.poi.ooxml;
	
	opens application to javafx.graphics, javafx.fxml;
}
