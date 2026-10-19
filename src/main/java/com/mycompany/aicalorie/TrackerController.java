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

// 3. CONTROLLER: Coordonează interacțiunea dintre Model și View
// ============================================================================

class TrackerController {
    private final TrackerModel model;
    private final TrackerView view;

    public TrackerController(TrackerModel model, TrackerView view) {
        this.model = model;
        this.view = view;

        initCustomRenderers();
        bindEvents();
        refreshAllViews();
    }

    private void initCustomRenderers() {
        view.setHistoryDayListCellRenderer(new DefaultListCellRenderer() {
            @Override
            public Component getListCellRendererComponent(JList<?> list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                JLabel label = (JLabel) super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                String dateKey = (String) value;
                DaySummary summary = model.getDaySummary(dateKey);

                String statusTag;
                if (summary.isCalorieMarginExceeded()) {
                    statusTag = " [WARNING: CALORIE MARGIN]";
                } else if (summary.isProteinNotReached) {
                    statusTag = " [WARNING: LOW PROTEIN]";
                } else {
                    statusTag = " [OK]";
                }

                label.setText(String.format("Day: %s (%.0f kcal)%s", dateKey, summary.totalCalories, statusTag));
                label.setFont(new Font("SansSerif", Font.BOLD, 12));
                label.setBorder(new EmptyBorder(6, 10, 6, 10));

                if (summary.hasAnyWarning()) {
                    if (isSelected) {
                        label.setBackground(TrackerView.BTN_DANGER);
                        label.setForeground(Color.WHITE);
                    } else {
                        label.setBackground(TrackerView.ALERT_RED_BG);
                        label.setForeground(TrackerView.BTN_DANGER);
                    }
                } else {
                    if (isSelected) {
                        label.setBackground(TrackerView.BTN_PRIMARY);
                        label.setForeground(Color.WHITE);
                    } else {
                        label.setBackground(TrackerView.ALERT_GREEN_BG);
                        label.setForeground(new Color(21, 128, 61));
                    }
                }
                return label;
            }
        });
    }

    private void bindEvents() {
        view.onDashboardNavClicked(e -> {
            refreshDashboard();
            view.showCard("DASHBOARD");
        });

        view.onHistoryNavClicked(e -> {
            refreshHistoryView();
            view.showCard("HISTORY");
        });

        view.onChangeDateClicked(e -> handlePromptChangeDate());
        view.onSetCalGoalClicked(e -> handlePromptSetCalorieGoal());
        view.onSetProtGoalClicked(e -> handlePromptSetProteinGoal());
        view.onSetMarginClicked(e -> handlePromptSetMargins());
        view.onAddMealClicked(e -> handleOpenAddMealDialog());
        view.onRemoveEntryClicked(e -> handleRemoveSelectedEntry());

        view.onFilterExceededToggled(e -> refreshHistoryView());

        view.onHistoryDaySelected(e -> {
            if (!e.getValueIsAdjusting()) {
                updateSelectedHistoryDayDetails(view.getSelectedHistoryDay());
            }
        });

        view.onSwitchToDayClicked(e -> {
            String selected = view.getSelectedHistoryDay();
            if (selected != null) {
                model.setCurrentDate(selected);
                refreshDashboard();
                view.showCard("DASHBOARD");
            }
        });
    }

    public void refreshAllViews() {
        refreshDashboard();
        refreshHistoryView();
    }

    private void refreshDashboard() {
        String currentDate = model.getCurrentDate();
        DaySummary summary = model.getDaySummary(currentDate);
        List<MealEntry> entries = model.getMealsForDate(currentDate);

        view.updateDashboard(
                currentDate,
                model.getCalorieGoal(),
                model.getProteinGoal(),
                model.getUnderMargin(),
                model.getOverMargin(),
                summary,
                entries
        );
    }

    private void refreshHistoryView() {
        boolean onlyExceeded = view.isFilterExceededSelected();
        List<String> dates = model.getHistoryDates(onlyExceeded);
        view.updateHistoryDaysList(dates);
        updateSelectedHistoryDayDetails(view.getSelectedHistoryDay());
    }

    private void updateSelectedHistoryDayDetails(String day) {
        if (day == null) {
            view.updateHistoryDayDetails(null, Collections.emptyList());
        } else {
            DaySummary summary = model.getDaySummary(day);
            List<MealEntry> entries = model.getMealsForDate(day);
            view.updateHistoryDayDetails(summary, entries);
        }
    }

    // --- Handlere pentru Dialoguri și Acțiuni ---
    private void handlePromptSetCalorieGoal() {
        String input = JOptionPane.showInputDialog(view.getMainFrame(),
                "Enter your daily Calorie Goal (kcal):", String.valueOf((int) model.getCalorieGoal()));
        if (input != null) {
            try {
                double val = Double.parseDouble(input.trim());
                if (val > 0) {
                    model.setCalorieGoal(val);
                    refreshAllViews();
                }
            } catch (NumberFormatException ex) {
                JOptionPane.showMessageDialog(view.getMainFrame(), "Invalid number entered.");
            }
        }
    }

    private void handlePromptSetProteinGoal() {
        String input = JOptionPane.showInputDialog(view.getMainFrame(),
                "Enter your daily Protein Goal (grams):", String.valueOf((int) model.getProteinGoal()));
        if (input != null) {
            try {
                double val = Double.parseDouble(input.trim());
                if (val > 0) {
                    model.setProteinGoal(val);
                    refreshAllViews();
                }
            } catch (NumberFormatException ex) {
                JOptionPane.showMessageDialog(view.getMainFrame(), "Invalid number entered.");
            }
        }
    }

    private void handlePromptSetMargins() {
        JTextField txtUnder = new JTextField(String.valueOf((int) model.getUnderMargin()), 8);
        JTextField txtOver = new JTextField(String.valueOf((int) model.getOverMargin()), 8);

        JPanel panel = new JPanel(new GridLayout(2, 2, 10, 10));
        panel.add(new JLabel("Under Margin (- kcal allowed):"));
        panel.add(txtUnder);
        panel.add(new JLabel("Over Margin (+ kcal allowed):"));
        panel.add(txtOver);

        int result = JOptionPane.showConfirmDialog(view.getMainFrame(), panel,
                "Set Under / Over Calorie Margin", JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);

        if (result == JOptionPane.OK_OPTION) {
            try {
                double u = Double.parseDouble(txtUnder.getText().trim());
                double o = Double.parseDouble(txtOver.getText().trim());
                if (u >= 0 && o >= 0) {
                    model.setMargins(u, o);
                    refreshAllViews();
                }
            } catch (NumberFormatException ex) {
                JOptionPane.showMessageDialog(view.getMainFrame(), "Please enter valid positive numbers for margins.");
            }
        }
    }

    private void handleRemoveSelectedEntry() {
        int selectedRow = view.getSelectedDashboardRow();
        if (selectedRow < 0) {
            JOptionPane.showMessageDialog(view.getMainFrame(), "Please select an entry from the table first to remove it.");
            return;
        }
        model.removeMealFromCurrentDate(selectedRow);
        refreshAllViews();
    }

    private void handlePromptChangeDate() {
        String input = JOptionPane.showInputDialog(view.getMainFrame(),
                "Enter Day Date (Format YYYY-MM-DD):", model.getCurrentDate());
        if (input != null && !input.trim().isEmpty()) {
            model.setCurrentDate(input.trim());
            refreshAllViews();
        }
    }

    private void handleOpenAddMealDialog() {
        JDialog dialog = new JDialog(view.getMainFrame(), "Add Meal to Day: " + model.getCurrentDate(), true);
        dialog.setSize(720, 620);
        dialog.setLocationRelativeTo(view.getMainFrame());
        dialog.getContentPane().setBackground(TrackerView.BG_MAIN);
        dialog.setLayout(new BorderLayout(12, 12));

        JPanel content = new JPanel();
        content.setLayout(new BoxLayout(content, BoxLayout.Y_AXIS));
        content.setBackground(TrackerView.BG_MAIN);
        content.setBorder(new EmptyBorder(18, 20, 18, 20));

        // 1. Selectare categorie masă
        JPanel mealTypePanel = TrackerView.createStyledCard("1. SELECT MEAL CATEGORY");
        JPanel radioGroupPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 20, 4));
        radioGroupPanel.setOpaque(false);

        ButtonGroup mealGroup = new ButtonGroup();
        JRadioButton rbBreakfast = new JRadioButton("Breakfast", true);
        JRadioButton rbLunch = new JRadioButton("Lunch");
        JRadioButton rbDinner = new JRadioButton("Dinner");
        JRadioButton rbSnack = new JRadioButton("Snack");

        for (JRadioButton rb : new JRadioButton[]{rbBreakfast, rbLunch, rbDinner, rbSnack}) {
            rb.setFont(new Font("SansSerif", Font.BOLD, 14));
            rb.setOpaque(false);
            rb.setForeground(TrackerView.TEXT_DARK);
            mealGroup.add(rb);
            radioGroupPanel.add(rb);
        }
        mealTypePanel.add(radioGroupPanel);

        // 2. Căutare în baza de date prin Model
        JPanel searchPanel = TrackerView.createStyledCard("2. SEARCH FOOD IN EXCEL DATABASE (e.g., chicken, beef, apple)");
        JPanel searchBar = new JPanel(new BorderLayout(10, 0));
        searchBar.setOpaque(false);

        JTextField txtSearch = new JTextField();
        txtSearch.setFont(new Font("SansSerif", Font.PLAIN, 15));
        txtSearch.setBorder(BorderFactory.createCompoundBorder(
                new LineBorder(new Color(148, 163, 184), 1),
                new EmptyBorder(6, 10, 6, 10)
        ));

        JButton btnSearch = TrackerView.createTangibleButton("Search Food", TrackerView.BTN_PRIMARY, Color.WHITE, 13);
        searchBar.add(txtSearch, BorderLayout.CENTER);
        searchBar.add(btnSearch, BorderLayout.EAST);

        DefaultListModel<Food> foodListModel = new DefaultListModel<>();
        JList<Food> foodJList = new JList<>(foodListModel);
        foodJList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        foodJList.setFont(new Font("SansSerif", Font.PLAIN, 13));
        foodJList.setFixedCellHeight(28);

        JScrollPane listScroll = new JScrollPane(foodJList);
        listScroll.setPreferredSize(new Dimension(640, 190));
        listScroll.setBorder(new LineBorder(new Color(203, 213, 225), 1));

        searchPanel.add(Box.createVerticalStrut(6));
        searchPanel.add(searchBar);
        searchPanel.add(Box.createVerticalStrut(10));
        searchPanel.add(listScroll);

        // 3. Introducere Gramaj & Previzualizare
        JPanel gramsPanel = TrackerView.createStyledCard("3. ENTER WEIGHT IN GRAMS & PREVIEW OUTPUT");
        JPanel gramsInputRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 12, 4));
        gramsInputRow.setOpaque(false);

        JLabel lblGrams = new JLabel("Amount (grams):");
        lblGrams.setFont(new Font("SansSerif", Font.BOLD, 14));
        JTextField txtGrams = new JTextField("150", 10);
        txtGrams.setFont(new Font("SansSerif", Font.BOLD, 15));
        txtGrams.setBorder(BorderFactory.createCompoundBorder(
                new LineBorder(new Color(148, 163, 184), 1),
                new EmptyBorder(5, 8, 5, 8)
        ));

        gramsInputRow.add(lblGrams);
        gramsInputRow.add(txtGrams);

        JLabel lblPreview = new JLabel("Output Preview: Select a food and enter grams...");
        lblPreview.setFont(new Font("SansSerif", Font.BOLD, 13));
        lblPreview.setForeground(TrackerView.BTN_PRIMARY);

        gramsPanel.add(gramsInputRow);
        gramsPanel.add(Box.createVerticalStrut(6));
        gramsPanel.add(lblPreview);

        Runnable performSearch = () -> {
            String query = txtSearch.getText();
            foodListModel.clear();
            List<Food> matches = model.searchFoods(query);
            for (Food f : matches) {
                foodListModel.addElement(f);
            }
            if (!foodListModel.isEmpty()) {
                foodJList.setSelectedIndex(0);
            } else {
                lblPreview.setText("No matching food found for: " + query);
            }
        };

        Runnable updatePreview = () -> {
            Food selected = foodJList.getSelectedValue();
            if (selected == null) return;
            try {
                double g = Double.parseDouble(txtGrams.getText().trim());
                if (g <= 0) {
                    lblPreview.setText("Enter an amount greater than 0g.");
                    return;
                }
                MealEntry tempEntry = new MealEntry(model.getCurrentDate(), "Preview", selected, g);
                lblPreview.setText(String.format(
                        "Output: %.1f kcal | Fats: %.1fg | Proteins: %.1fg | Fiber: %.1fg | Carbs: %.1fg | Sugars: %.1fg",
                        tempEntry.getCalories(), tempEntry.getFats(), tempEntry.getProteins(),
                        tempEntry.getFiber(), tempEntry.getCarbs(), tempEntry.getSugars()
                ));
            } catch (NumberFormatException ex) {
                lblPreview.setText("Invalid number for grams.");
            }
        };

        btnSearch.addActionListener(e -> performSearch.run());
        txtSearch.addActionListener(e -> performSearch.run());
        foodJList.addListSelectionListener(e -> updatePreview.run());
        txtGrams.getDocument().addDocumentListener(new DocumentListener() {
            public void insertUpdate(DocumentEvent e) { updatePreview.run(); }
            public void removeUpdate(DocumentEvent e) { updatePreview.run(); }
            public void changedUpdate(DocumentEvent e) { updatePreview.run(); }
        });

        performSearch.run();

        // 4. Butoane de confirmare
        JPanel bottomBtns = new JPanel(new FlowLayout(FlowLayout.RIGHT, 12, 10));
        bottomBtns.setOpaque(false);

        JButton btnCancel = TrackerView.createTangibleButton("Cancel", TrackerView.BTN_DARK, Color.WHITE, 14);
        JButton btnConfirmAdd = TrackerView.createTangibleButton("Confirm & Add Meal", TrackerView.BTN_SUCCESS, Color.WHITE, 14);

        btnCancel.addActionListener(e -> dialog.dispose());
        btnConfirmAdd.addActionListener(e -> {
            Food selectedFood = foodJList.getSelectedValue();
            if (selectedFood == null) {
                JOptionPane.showMessageDialog(dialog, "Please select a food from the list first!");
                return;
            }
            double grams;
            try {
                grams = Double.parseDouble(txtGrams.getText().trim());
                if (grams <= 0) throw new NumberFormatException();
            } catch (NumberFormatException ex) {
                JOptionPane.showMessageDialog(dialog, "Please enter a valid positive number for grams!");
                return;
            }

            String mealType = "Breakfast";
            if (rbLunch.isSelected()) mealType = "Lunch";
            else if (rbDinner.isSelected()) mealType = "Dinner";
            else if (rbSnack.isSelected()) mealType = "Snack";

            model.addMealToCurrentDate(mealType, selectedFood, grams);
            refreshAllViews();
            dialog.dispose();
        });

        bottomBtns.add(btnCancel);
        bottomBtns.add(btnConfirmAdd);

        content.add(mealTypePanel);
        content.add(Box.createVerticalStrut(12));
        content.add(searchPanel);
        content.add(Box.createVerticalStrut(12));
        content.add(gramsPanel);

        dialog.add(content, BorderLayout.CENTER);
        dialog.add(bottomBtns, BorderLayout.SOUTH);
        dialog.setVisible(true);
    }
}

// ============================================================================
