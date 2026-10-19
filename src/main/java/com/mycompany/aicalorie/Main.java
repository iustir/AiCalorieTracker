import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.border.EmptyBorder;
import javax.swing.border.LineBorder;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.JTableHeader;
import java.awt.*;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.*;
import java.time.LocalDate;
import java.util.*;
import java.util.List;

// 4. MAIN: Punctul de intrare care asamblează Model, View și Controller
// ============================================================================

public class Main {
    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getCrossPlatformLookAndFeelClassName());
        } catch (Exception ignored) {}

        // 1. Inițializăm Modelul și încărcăm datele
        System.out.println("Loading food database from Excel... Please wait.");
        TrackerModel model = new TrackerModel();

        if (!model.loadDatabase()) {
            JOptionPane.showMessageDialog(null,
                    "Error: Database is empty or could not be found at:\n" + model.getFilePath(),
                    "Database Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        model.loadSavedHistory();

        // 2. Pornim View-ul și Controllerul pe firul de execuție Swing (EDT)
        SwingUtilities.invokeLater(() -> {
            TrackerView view = new TrackerView();
            view.initGUI();
            new TrackerController(model, view);
            view.showWindow();
        });
    }
}
