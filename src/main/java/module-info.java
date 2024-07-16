module org.example.loadplan_automation {
    requires javafx.controls;
    requires javafx.fxml;

    requires com.dlsc.formsfx;
    requires org.json;
    requires org.apache.poi.poi;

    opens org.example.loadplan_automation to javafx.fxml;
    exports org.example.loadplan_automation;
}