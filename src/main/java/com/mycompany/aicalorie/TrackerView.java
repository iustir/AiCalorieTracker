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

// 2. VIEW: Interfața grafică Swing (Doar afișare și captare evenimente)
// ============================================================================

class TrackerView {
    // Paleta de culori Premium
    public static final Color BG_MAIN = new Color(241, 245, 249);
    public static final Color BG_CARD = new Color(255, 255, 255);
    public static final Color BG_HEADER = new Color(15, 23, 42);
    public static final Color TEXT_DARK = new Color(30, 41, 59);
    public static final Color TEXT_MUTED = new Color(100, 116, 139);
    public static final Color BTN_PRIMARY = new Color(37, 99, 235);
    public static final Color BTN_SUCCESS = new Color(22, 163, 74);
    public static final Color BTN_WARNING = new Color(217, 119, 6);
    public static final Color BTN_PURPLE = new Color(124, 58, 237);
    public static final Color BTN_DANGER = new Color(220, 38, 38);
    public static final Color BTN_DARK = new Color(51, 65, 85);
    public static final Color ALERT_RED_BG = new Color(254, 226, 226);
    public static final Color ALERT_RED_BORDER = new Color(220, 38, 38);
    public static final Color ALERT_GREEN_BG = new Color(220, 252, 231);
    public static final Color ALERT_GREEN_BORDER = new Color(22, 163, 74);

    // Componente principale
    private JFrame mainFrame;
    private CardLayout cardLayout;
    private JPanel contentPanel;

    // Butoane Navigare
    private JButton btnDashboard;
    private JButton btnHistory;
    private JButton btnChangeDate;
    private JLabel lblActiveDate;

    // Componente Dashboard
    private JLabel lblCalorieSummary;
    private JLabel lblMarginSummary;
    private JLabel lblProteinSummary;
    private JLabel lblMacrosSummary;
    private JPanel warningBannerPanel;
    private JLabel lblCalorieWarning;
    private JLabel lblProteinWarning;
    private DefaultTableModel dashboardTableModel;
    private JTable dashboardTable;

    private JButton btnSetCalGoal;
    private JButton btnSetProtGoal;
    private JButton btnSetMargin;
    private JButton btnAddMeal;
    private JButton btnRemoveEntry;

    // Componente History
    private DefaultListModel<String> historyDaysListModel;
    private JList<String> historyDaysList;
    private JCheckBox chkFilterExceeded;
    private JButton btnSwitchToDay;
    private JLabel lblHistoryDayTitle;
    private JLabel lblHistoryDayWarnings;
    private DefaultTableModel historyTableModel;
    private JTable historyTable;

    public void initGUI() {
        mainFrame = new JFrame("Calorie Tracker Pro - Nutritional Dashboard & History (MVC)");
        mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        mainFrame.setSize(1220, 820);
        mainFrame.setMinimumSize(new Dimension(1050, 720));
        mainFrame.setLocationRelativeTo(null);
        mainFrame.getContentPane().setBackground(BG_MAIN);
        mainFrame.setLayout(new BorderLayout());

        mainFrame.add(createTopNavBar(), BorderLayout.NORTH);

        cardLayout = new CardLayout();
        contentPanel = new JPanel(cardLayout);
        contentPanel.setBackground(BG_MAIN);

        contentPanel.add(createDashboardPanel(), "DASHBOARD");
        contentPanel.add(createHistoryPanel(), "HISTORY");

        mainFrame.add(contentPanel, BorderLayout.CENTER);
    }

    public void showWindow() {
        mainFrame.setVisible(true);
    }

    public JFrame getMainFrame() {
        return mainFrame;
    }

    public void showCard(String cardName) {
        cardLayout.show(contentPanel, cardName);
    }

    // --- Construire Panouri ---
    private JPanel createTopNavBar() {
        JPanel navBar = new JPanel(new BorderLayout());
        navBar.setBackground(BG_HEADER);
        navBar.setBorder(new EmptyBorder(14, 24, 14, 24));

        JLabel title = new JLabel("CALORIE TRACKER AI");
        title.setFont(new Font("SansSerif", Font.BOLD, 22));
        title.setForeground(Color.WHITE);

        JPanel navButtons = new JPanel(new FlowLayout(FlowLayout.CENTER, 15, 0));
        navButtons.setOpaque(false);

        btnDashboard = createTangibleButton("Dashboard", BTN_PRIMARY, Color.WHITE, 14);
        btnHistory = createTangibleButton("History", BTN_PURPLE, Color.WHITE, 14);

        navButtons.add(btnDashboard);
        navButtons.add(btnHistory);

        JPanel datePanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 0));
        datePanel.setOpaque(false);

        lblActiveDate = new JLabel("Active Day: ");
        lblActiveDate.setFont(new Font("SansSerif", Font.BOLD, 14));
        lblActiveDate.setForeground(new Color(226, 232, 240));

        btnChangeDate = createTangibleButton("Change Day", BTN_DARK, Color.WHITE, 12);

        datePanel.add(lblActiveDate);
        datePanel.add(btnChangeDate);

        navBar.add(title, BorderLayout.WEST);
        navBar.add(navButtons, BorderLayout.CENTER);
        navBar.add(datePanel, BorderLayout.EAST);

        return navBar;
    }

    private JPanel createDashboardPanel() {
        JPanel dashboard = new JPanel(new BorderLayout(0, 16));
        dashboard.setBackground(BG_MAIN);
        dashboard.setBorder(new EmptyBorder(20, 24, 20, 24));

        JPanel topContainer = new JPanel();
        topContainer.setLayout(new BoxLayout(topContainer, BoxLayout.Y_AXIS));
        topContainer.setOpaque(false);

        JPanel cardsGrid = new JPanel(new GridLayout(1, 3, 16, 0));
        cardsGrid.setOpaque(false);

        JPanel calCard = createStyledCard("CALORIE GOAL & STATUS");
        lblCalorieSummary = new JLabel("0.00 / 2000.00 kcal");
        lblCalorieSummary.setFont(new Font("SansSerif", Font.BOLD, 20));
        lblCalorieSummary.setForeground(TEXT_DARK);
        lblMarginSummary = new JLabel("Allowed Margin: 1500 - 2500 kcal (-500 / +500)");
        lblMarginSummary.setFont(new Font("SansSerif", Font.PLAIN, 12));
        lblMarginSummary.setForeground(TEXT_MUTED);
        calCard.add(lblCalorieSummary);
        calCard.add(Box.createVerticalStrut(6));
        calCard.add(lblMarginSummary);

        JPanel protCard = createStyledCard("PROTEIN GOAL & STATUS");
        lblProteinSummary = new JLabel("0.00 / 150.00 g");
        lblProteinSummary.setFont(new Font("SansSerif", Font.BOLD, 20));
        lblProteinSummary.setForeground(TEXT_DARK);
        JLabel protHint = new JLabel("Daily target minimum requirement");
        protHint.setFont(new Font("SansSerif", Font.PLAIN, 12));
        protHint.setForeground(TEXT_MUTED);
        protCard.add(lblProteinSummary);
        protCard.add(Box.createVerticalStrut(6));
        protCard.add(protHint);

        JPanel macroCard = createStyledCard("OTHER NUTRITIONAL TOTALS");
        lblMacrosSummary = new JLabel("Fats: 0g | Carbs: 0g | Fiber: 0g | Sugars: 0g");
        lblMacrosSummary.setFont(new Font("SansSerif", Font.BOLD, 13));
        lblMacrosSummary.setForeground(TEXT_DARK);
        JLabel macroSub = new JLabel("Calculated from USDA Foundation Foods DB");
        macroSub.setFont(new Font("SansSerif", Font.PLAIN, 12));
        macroSub.setForeground(TEXT_MUTED);
        macroCard.add(lblMacrosSummary);
        macroCard.add(Box.createVerticalStrut(6));
        macroCard.add(macroSub);

        cardsGrid.add(calCard);
        cardsGrid.add(protCard);
        cardsGrid.add(macroCard);

        warningBannerPanel = new JPanel(new GridLayout(2, 1, 0, 4));
        warningBannerPanel.setBackground(ALERT_RED_BG);
        warningBannerPanel.setBorder(BorderFactory.createCompoundBorder(
                new LineBorder(ALERT_RED_BORDER, 2),
                new EmptyBorder(10, 16, 10, 16)
        ));

        lblCalorieWarning = new JLabel("Calorie status");
        lblCalorieWarning.setFont(new Font("SansSerif", Font.BOLD, 14));
        lblProteinWarning = new JLabel("Warning protein goal not reached!");
        lblProteinWarning.setFont(new Font("SansSerif", Font.BOLD, 14));

        warningBannerPanel.add(lblCalorieWarning);
        warningBannerPanel.add(lblProteinWarning);

        JPanel actionToolbar = new JPanel(new FlowLayout(FlowLayout.LEFT, 12, 10));
        actionToolbar.setBackground(BG_CARD);
        actionToolbar.setBorder(BorderFactory.createCompoundBorder(
                new LineBorder(new Color(203, 213, 225), 1),
                new EmptyBorder(6, 10, 6, 10)
        ));

        btnSetCalGoal = createTangibleButton("Set Calorie Goal", BTN_PRIMARY, Color.WHITE, 13);
        btnSetProtGoal = createTangibleButton("Set Protein Goal", BTN_PRIMARY, Color.WHITE, 13);
        btnSetMargin = createTangibleButton("Set Under/Over Margin", BTN_WARNING, Color.WHITE, 13);
        btnAddMeal = createTangibleButton("+ Add Meal (Breakfast / Lunch / Dinner / Snack)", BTN_SUCCESS, Color.WHITE, 13);
        btnRemoveEntry = createTangibleButton("Remove Selected Entry", BTN_DANGER, Color.WHITE, 13);

        actionToolbar.add(btnSetCalGoal);
        actionToolbar.add(btnSetProtGoal);
        actionToolbar.add(btnSetMargin);
        actionToolbar.add(btnAddMeal);
        actionToolbar.add(btnRemoveEntry);

        topContainer.add(cardsGrid);
        topContainer.add(Box.createVerticalStrut(14));
        topContainer.add(warningBannerPanel);
        topContainer.add(Box.createVerticalStrut(14));
        topContainer.add(actionToolbar);

        dashboard.add(topContainer, BorderLayout.NORTH);

        String[] columns = {
                "Date", "Meal Type", "Food Name", "Grams (g)",
                "Calories (kcal)", "Fats (g)", "Proteins (g)", "Fiber (g)", "Carbs (g)", "Sugars (g)"
        };
        dashboardTableModel = new DefaultTableModel(columns, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        dashboardTable = createStyledTable(dashboardTableModel);
        JScrollPane scrollPane = new JScrollPane(dashboardTable);
        scrollPane.setBorder(new LineBorder(new Color(203, 213, 225), 1));
        scrollPane.getViewport().setBackground(Color.WHITE);

        dashboard.add(scrollPane, BorderLayout.CENTER);

        return dashboard;
    }

    private JPanel createHistoryPanel() {
        JPanel historyPanel = new JPanel(new BorderLayout(18, 0));
        historyPanel.setBackground(BG_MAIN);
        historyPanel.setBorder(new EmptyBorder(20, 24, 20, 24));

        JPanel leftPanel = new JPanel(new BorderLayout(0, 10));
        leftPanel.setBackground(BG_CARD);
        leftPanel.setPreferredSize(new Dimension(360, 0));
        leftPanel.setBorder(BorderFactory.createCompoundBorder(
                new LineBorder(new Color(203, 213, 225), 1),
                new EmptyBorder(14, 14, 14, 14)
        ));

        JLabel lblLeftTitle = new JLabel("HISTORY PATHWAY: SELECT DAY");
        lblLeftTitle.setFont(new Font("SansSerif", Font.BOLD, 14));
        lblLeftTitle.setForeground(TEXT_DARK);

        chkFilterExceeded = new JCheckBox("Show ONLY days with Under/Over Warning");
        chkFilterExceeded.setBackground(BG_CARD);
        chkFilterExceeded.setForeground(BTN_DANGER);
        chkFilterExceeded.setFont(new Font("SansSerif", Font.BOLD, 12));

        JPanel leftTop = new JPanel(new GridLayout(2, 1, 0, 6));
        leftTop.setOpaque(false);
        leftTop.add(lblLeftTitle);
        leftTop.add(chkFilterExceeded);

        historyDaysListModel = new DefaultListModel<>();
        historyDaysList = new JList<>(historyDaysListModel);
        historyDaysList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        historyDaysList.setFixedCellHeight(48);

        JScrollPane dayListScroll = new JScrollPane(historyDaysList);
        dayListScroll.setBorder(new LineBorder(new Color(203, 213, 225), 1));

        btnSwitchToDay = createTangibleButton("Edit Selected Day in Dashboard", BTN_PRIMARY, Color.WHITE, 12);

        leftPanel.add(leftTop, BorderLayout.NORTH);
        leftPanel.add(dayListScroll, BorderLayout.CENTER);
        leftPanel.add(btnSwitchToDay, BorderLayout.SOUTH);

        JPanel rightPanel = new JPanel(new BorderLayout(0, 12));
        rightPanel.setOpaque(false);

        JPanel rightHeader = createStyledCard("SELECTED DAY DETAILS");
        lblHistoryDayTitle = new JLabel("Select a day from the list");
        lblHistoryDayTitle.setFont(new Font("SansSerif", Font.BOLD, 18));
        lblHistoryDayTitle.setForeground(TEXT_DARK);

        lblHistoryDayWarnings = new JLabel(" ");
        lblHistoryDayWarnings.setFont(new Font("SansSerif", Font.BOLD, 13));
        lblHistoryDayWarnings.setForeground(BTN_DANGER);

        rightHeader.add(lblHistoryDayTitle);
        rightHeader.add(Box.createVerticalStrut(6));
        rightHeader.add(lblHistoryDayWarnings);

        String[] columns = {
                "Date", "Meal Type", "Food Name", "Grams (g)",
                "Calories (kcal)", "Fats (g)", "Proteins (g)", "Fiber (g)", "Carbs (g)", "Sugars (g)"
        };
        historyTableModel = new DefaultTableModel(columns, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        historyTable = createStyledTable(historyTableModel);
        JScrollPane rightTableScroll = new JScrollPane(historyTable);
        rightTableScroll.setBorder(new LineBorder(new Color(203, 213, 225), 1));
        rightTableScroll.getViewport().setBackground(Color.WHITE);

        rightPanel.add(rightHeader, BorderLayout.NORTH);
        rightPanel.add(rightTableScroll, BorderLayout.CENTER);

        historyPanel.add(leftPanel, BorderLayout.WEST);
        historyPanel.add(rightPanel, BorderLayout.CENTER);

        return historyPanel;
    }

    // --- Metode de Actualizare UI (Apelate de Controller) ---
    public void updateDashboard(String currentDate, double calorieGoal, double proteinGoal,
                                double underMargin, double overMargin,
                                DaySummary summary, List<MealEntry> entries) {
        lblActiveDate.setText("Active Day: " + currentDate);
        dashboardTableModel.setRowCount(0);

        for (MealEntry entry : entries) {
            dashboardTableModel.addRow(new Object[]{
                    entry.getDate(),
                    entry.getMealType(),
                    entry.getFood().getName(),
                    String.format("%.1f g", entry.getGrams()),
                    String.format("%.2f", entry.getCalories()),
                    String.format("%.2f", entry.getFats()),
                    String.format("%.2f", entry.getProteins()),
                    String.format("%.2f", entry.getFiber()),
                    String.format("%.2f", entry.getCarbs()),
                    String.format("%.2f", entry.getSugars())
            });
        }

        lblCalorieSummary.setText(String.format("%.2f / %.0f kcal", summary.totalCalories, calorieGoal));
        lblMarginSummary.setText(String.format("Allowed Margin: %.0f - %.0f kcal (-%.0f / +%.0f)",
                summary.minAllowedCalories, summary.maxAllowedCalories, underMargin, overMargin));
        lblProteinSummary.setText(String.format("%.2f / %.0f g", summary.totalProteins, proteinGoal));
        lblMacrosSummary.setText(String.format("Fats: %.1fg | Carbs: %.1fg | Fiber: %.1fg | Sugars: %.1fg",
                summary.totalFats, summary.totalCarbs, summary.totalFiber, summary.totalSugars));

        if (summary.isCalorieUnder) {
            lblCalorieWarning.setText(String.format(
                    "WARNING: You are UNDER your calorie margin by %.1f kcal! (Minimum: %.0f kcal)",
                    (summary.minAllowedCalories - summary.totalCalories), summary.minAllowedCalories));
            lblCalorieWarning.setForeground(BTN_DANGER);
        } else if (summary.isCalorieOver) {
            lblCalorieWarning.setText(String.format(
                    "WARNING: You are OVER your calorie margin by +%.1f kcal! (Maximum: %.0f kcal)",
                    (summary.totalCalories - summary.maxAllowedCalories), summary.maxAllowedCalories));
            lblCalorieWarning.setForeground(BTN_DANGER);
        } else {
            lblCalorieWarning.setText("Calories are within your allowed Under/Over margin.");
            lblCalorieWarning.setForeground(new Color(21, 128, 61));
        }

        if (summary.isProteinNotReached) {
            lblProteinWarning.setText(String.format(
                    "Warning protein goal not reached! (Current: %.1fg / Goal: %.0fg)",
                    summary.totalProteins, proteinGoal));
            lblProteinWarning.setForeground(BTN_DANGER);
        } else {
            lblProteinWarning.setText("Protein goal reached successfully!");
            lblProteinWarning.setForeground(new Color(21, 128, 61));
        }

        if (summary.hasAnyWarning()) {
            warningBannerPanel.setBackground(ALERT_RED_BG);
            warningBannerPanel.setBorder(BorderFactory.createCompoundBorder(
                    new LineBorder(ALERT_RED_BORDER, 2),
                    new EmptyBorder(10, 16, 10, 16)
            ));
        } else {
            warningBannerPanel.setBackground(ALERT_GREEN_BG);
            warningBannerPanel.setBorder(BorderFactory.createCompoundBorder(
                    new LineBorder(ALERT_GREEN_BORDER, 2),
                    new EmptyBorder(10, 16, 10, 16)
            ));
        }
    }

    public void setHistoryDayListCellRenderer(ListCellRenderer<? super String> renderer) {
        historyDaysList.setCellRenderer(renderer);
    }

    public void updateHistoryDaysList(List<String> days) {
        String previouslySelected = historyDaysList.getSelectedValue();
        historyDaysListModel.clear();
        for (String day : days) {
            historyDaysListModel.addElement(day);
        }

        if (!historyDaysListModel.isEmpty()) {
            if (previouslySelected != null && historyDaysListModel.contains(previouslySelected)) {
                historyDaysList.setSelectedValue(previouslySelected, true);
            } else {
                historyDaysList.setSelectedIndex(0);
            }
        }
    }

    public void updateHistoryDayDetails(DaySummary summary, List<MealEntry> entries) {
        historyTableModel.setRowCount(0);
        if (summary == null) {
            lblHistoryDayTitle.setText("No day selected");
            lblHistoryDayWarnings.setText(" ");
            return;
        }

        for (MealEntry entry : entries) {
            historyTableModel.addRow(new Object[]{
                    entry.getDate(),
                    entry.getMealType(),
                    entry.getFood().getName(),
                    String.format("%.1f g", entry.getGrams()),
                    String.format("%.2f", entry.getCalories()),
                    String.format("%.2f", entry.getFats()),
                    String.format("%.2f", entry.getProteins()),
                    String.format("%.2f", entry.getFiber()),
                    String.format("%.2f", entry.getCarbs()),
                    String.format("%.2f", entry.getSugars())
            });
        }

        lblHistoryDayTitle.setText(String.format("Day: %s  |  Total Calories: %.1f kcal  |  Total Proteins: %.1f g",
                summary.date, summary.totalCalories, summary.totalProteins));

        List<String> warnings = new ArrayList<>();
        if (summary.isCalorieUnder) {
            warnings.add(String.format("UNDER CALORIE MARGIN (-%.1f kcal)", summary.minAllowedCalories - summary.totalCalories));
        } else if (summary.isCalorieOver) {
            warnings.add(String.format("OVER CALORIE MARGIN (+%.1f kcal)", summary.totalCalories - summary.maxAllowedCalories));
        }
        if (summary.isProteinNotReached) {
            warnings.add("Warning protein goal not reached!");
        }

        if (warnings.isEmpty()) {
            lblHistoryDayWarnings.setText("Status: All goals & margins respected on this day.");
            lblHistoryDayWarnings.setForeground(new Color(21, 128, 61));
        } else {
            lblHistoryDayWarnings.setText(String.join("   |   ", warnings));
            lblHistoryDayWarnings.setForeground(BTN_DANGER);
        }
    }

    // --- Getters stare curentă View ---
    public int getSelectedDashboardRow() { return dashboardTable.getSelectedRow(); }
    public String getSelectedHistoryDay() { return historyDaysList.getSelectedValue(); }
    public boolean isFilterExceededSelected() { return chkFilterExceeded.isSelected(); }

    // --- Atașare Listeners pentru Controller ---
    public void onDashboardNavClicked(ActionListener l) { btnDashboard.addActionListener(l); }
    public void onHistoryNavClicked(ActionListener l) { btnHistory.addActionListener(l); }
    public void onChangeDateClicked(ActionListener l) { btnChangeDate.addActionListener(l); }
    public void onSetCalGoalClicked(ActionListener l) { btnSetCalGoal.addActionListener(l); }
    public void onSetProtGoalClicked(ActionListener l) { btnSetProtGoal.addActionListener(l); }
    public void onSetMarginClicked(ActionListener l) { btnSetMargin.addActionListener(l); }
    public void onAddMealClicked(ActionListener l) { btnAddMeal.addActionListener(l); }
    public void onRemoveEntryClicked(ActionListener l) { btnRemoveEntry.addActionListener(l); }
    public void onFilterExceededToggled(ActionListener l) { chkFilterExceeded.addActionListener(l); }
    public void onSwitchToDayClicked(ActionListener l) { btnSwitchToDay.addActionListener(l); }
    public void onHistoryDaySelected(javax.swing.event.ListSelectionListener l) { historyDaysList.addListSelectionListener(l); }

    // --- Helpere Design ---
    public static JButton createTangibleButton(String text, Color bg, Color fg, int fontSize) {
        JButton btn = new JButton(text);
        btn.setFont(new Font("SansSerif", Font.BOLD, fontSize));
        btn.setBackground(bg);
        btn.setForeground(fg);
        btn.setOpaque(true);
        btn.setContentAreaFilled(true);
        btn.setFocusPainted(false);
        btn.setCursor(new Cursor(Cursor.HAND_CURSOR));

        Border outer = new LineBorder(bg.darker(), 2);
        Border inner = new EmptyBorder(8, 16, 8, 16);
        btn.setBorder(BorderFactory.createCompoundBorder(outer, inner));

        btn.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) { btn.setBackground(bg.darker()); }
            @Override
            public void mouseExited(MouseEvent e) { btn.setBackground(bg); }
        });
        return btn;
    }

    public static JPanel createStyledCard(String titleText) {
        JPanel card = new JPanel();
        card.setLayout(new BoxLayout(card, BoxLayout.Y_AXIS));
        card.setBackground(BG_CARD);
        card.setBorder(BorderFactory.createCompoundBorder(
                new LineBorder(new Color(203, 213, 225), 1),
                new EmptyBorder(12, 16, 12, 16)
        ));

        JLabel lblTitle = new JLabel(titleText);
        lblTitle.setFont(new Font("SansSerif", Font.BOLD, 11));
        lblTitle.setForeground(TEXT_MUTED);
        card.add(lblTitle);
        card.add(Box.createVerticalStrut(6));
        return card;
    }

    private JTable createStyledTable(DefaultTableModel model) {
        JTable table = new JTable(model);
        table.setRowHeight(30);
        table.setFont(new Font("SansSerif", Font.PLAIN, 13));
        table.setSelectionBackground(new Color(219, 234, 254));
        table.setSelectionForeground(TEXT_DARK);
        table.setGridColor(new Color(226, 232, 240));

        JTableHeader header = table.getTableHeader();
        header.setBackground(BG_HEADER);
        header.setForeground(Color.WHITE);
        header.setFont(new Font("SansSerif", Font.BOLD, 13));
        header.setPreferredSize(new Dimension(0, 36));

        DefaultTableCellRenderer centerRenderer = new DefaultTableCellRenderer();
        centerRenderer.setHorizontalAlignment(JLabel.CENTER);
        for (int i = 0; i < table.getColumnCount(); i++) {
            if (i != 2) {
                table.getColumnModel().getColumn(i).setCellRenderer(centerRenderer);
            }
        }
        table.getColumnModel().getColumn(2).setPreferredWidth(260);
        return table;
    }
}

// ============================================================================
