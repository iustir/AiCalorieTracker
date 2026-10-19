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

// 1. MODEL: Entitățile de date și logica de business + persistență (Fără UI)
// ============================================================================

class Food {
    private final String name;
    private final double calories, proteins, fats, carbs, fiber, sugars;

    public Food(String name, double calories, double proteins, double fats, double carbs, double fiber, double sugars) {
        this.name = name;
        this.calories = calories;
        this.proteins = proteins;
        this.fats = fats;
        this.carbs = carbs;
        this.fiber = fiber;
        this.sugars = sugars;
    }

    public String getName() { return name; }
    public double getCalories() { return calories; }
    public double getProteins() { return proteins; }
    public double getFats() { return fats; }
    public double getCarbs() { return carbs; }
    public double getFiber() { return fiber; }
    public double getSugars() { return sugars; }

    @Override
    public String toString() {
        return String.format("%s (%.0f kcal, %.1fg P / 100g)", name, calories, proteins);
    }
}

class MealEntry {
    private final String date;
    private final String mealType; // Breakfast, Lunch, Dinner, Snack
    private final Food food;
    private final double grams;

    public MealEntry(String date, String mealType, Food food, double grams) {
        this.date = date;
        this.mealType = mealType;
        this.food = food;
        this.grams = grams;
    }

    public String getDate() { return date; }
    public String getMealType() { return mealType; }
    public Food getFood() { return food; }
    public double getGrams() { return grams; }

    public double getRatio() { return grams / 100.0; }
    public double getCalories() { return food.getCalories() * getRatio(); }
    public double getProteins() { return food.getProteins() * getRatio(); }
    public double getFats() { return food.getFats() * getRatio(); }
    public double getCarbs() { return food.getCarbs() * getRatio(); }
    public double getFiber() { return food.getFiber() * getRatio(); }
    public double getSugars() { return food.getSugars() * getRatio(); }
}

// Obiect sumar calculat de Model pentru o anumită zi
class DaySummary {
    public final String date;
    public final double totalCalories, totalFats, totalProteins, totalFiber, totalCarbs, totalSugars;
    public final double minAllowedCalories, maxAllowedCalories;
    public final boolean isCalorieUnder, isCalorieOver, isProteinNotReached;

    public DaySummary(String date, double totalCalories, double totalFats, double totalProteins,
                      double totalFiber, double totalCarbs, double totalSugars,
                      double minAllowedCalories, double maxAllowedCalories,
                      boolean isCalorieUnder, boolean isCalorieOver, boolean isProteinNotReached) {
        this.date = date;
        this.totalCalories = totalCalories;
        this.totalFats = totalFats;
        this.totalProteins = totalProteins;
        this.totalFiber = totalFiber;
        this.totalCarbs = totalCarbs;
        this.totalSugars = totalSugars;
        this.minAllowedCalories = minAllowedCalories;
        this.maxAllowedCalories = maxAllowedCalories;
        this.isCalorieUnder = isCalorieUnder;
        this.isCalorieOver = isCalorieOver;
        this.isProteinNotReached = isProteinNotReached;
    }

    public boolean isCalorieMarginExceeded() {
        return isCalorieUnder || isCalorieOver;
    }

    public boolean hasAnyWarning() {
        return isCalorieUnder || isCalorieOver || isProteinNotReached;
    }
}

class TrackerModel {
    private static final String FILE_PATH = "C:\\Users\\cipri\\OneDrive\\Documente\\NetBeansProjects\\Main\\USDA_Foundation_Foods_Complet (1).xlsx";
    private static final String SAVE_FILE = "meal_history_data.txt";

    private final List<Food> foodDatabase = new ArrayList<>();
    private final Map<String, List<MealEntry>> historyByDate = new TreeMap<>(Collections.reverseOrder());

    private double calorieGoal = 2000.0;
    private double proteinGoal = 150.0;
    private double underMargin = 500.0;
    private double overMargin = 500.0;
    private String currentDate = LocalDate.now().toString();

    public TrackerModel() {
        historyByDate.putIfAbsent(currentDate, new ArrayList<>());
    }

    // --- Getters & Setters pentru Starea Modelului ---
    public String getFilePath() { return FILE_PATH; }
    public List<Food> getFoodDatabase() { return Collections.unmodifiableList(foodDatabase); }
    public double getCalorieGoal() { return calorieGoal; }
    public double getProteinGoal() { return proteinGoal; }
    public double getUnderMargin() { return underMargin; }
    public double getOverMargin() { return overMargin; }
    public String getCurrentDate() { return currentDate; }

    public void setCalorieGoal(double calorieGoal) {
        this.calorieGoal = calorieGoal;
        saveHistory();
    }

    public void setProteinGoal(double proteinGoal) {
        this.proteinGoal = proteinGoal;
        saveHistory();
    }

    public void setMargins(double underMargin, double overMargin) {
        this.underMargin = underMargin;
        this.overMargin = overMargin;
        saveHistory();
    }

    public void setCurrentDate(String date) {
        this.currentDate = date;
        historyByDate.putIfAbsent(currentDate, new ArrayList<>());
        saveHistory();
    }

    // --- Operații de Business ---
    public List<Food> searchFoods(String query) {
        String cleanQuery = query == null ? "" : query.trim().toLowerCase();
        List<Food> results = new ArrayList<>();
        for (Food f : foodDatabase) {
            if (f.getName().toLowerCase().contains(cleanQuery)) {
                results.add(f);
            }
        }
        return results;
    }

    public void addMealToCurrentDate(String mealType, Food food, double grams) {
        historyByDate.putIfAbsent(currentDate, new ArrayList<>());
        historyByDate.get(currentDate).add(new MealEntry(currentDate, mealType, food, grams));
        saveHistory();
    }

    public void removeMealFromCurrentDate(int index) {
        List<MealEntry> entries = historyByDate.get(currentDate);
        if (entries != null && index >= 0 && index < entries.size()) {
            entries.remove(index);
            saveHistory();
        }
    }

    public List<MealEntry> getMealsForDate(String date) {
        if (date == null) return Collections.emptyList();
        return historyByDate.getOrDefault(date, Collections.emptyList());
    }

    public List<String> getHistoryDates(boolean onlyCalorieMarginExceeded) {
        List<String> dates = new ArrayList<>();
        for (String day : historyByDate.keySet()) {
            DaySummary summary = getDaySummary(day);
            if (onlyCalorieMarginExceeded && !summary.isCalorieMarginExceeded()) {
                continue;
            }
            dates.add(day);
        }
        return dates;
    }

    public DaySummary getDaySummary(String date) {
        List<MealEntry> entries = getMealsForDate(date);
        double totalCal = 0, totalFat = 0, totalProt = 0, totalFib = 0, totalCarb = 0, totalSug = 0;

        for (MealEntry entry : entries) {
            totalCal += entry.getCalories();
            totalFat += entry.getFats();
            totalProt += entry.getProteins();
            totalFib += entry.getFiber();
            totalCarb += entry.getCarbs();
            totalSug += entry.getSugars();
        }

        double minAllowed = calorieGoal - underMargin;
        double maxAllowed = calorieGoal + overMargin;

        boolean calUnder = totalCal < minAllowed;
        boolean calOver = totalCal > maxAllowed;
        boolean protNotReached = totalProt < proteinGoal;

        return new DaySummary(date, totalCal, totalFat, totalProt, totalFib, totalCarb, totalSug,
                minAllowed, maxAllowed, calUnder, calOver, protNotReached);
    }

    // --- Persistență: Citire Excel & Istoric Local ---
    public boolean loadDatabase() {
        foodDatabase.clear();
        try (FileInputStream fis = new FileInputStream(new File(FILE_PATH));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(1);
            for (int i = 2; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String name = getStringValue(row.getCell(1));
                if (name.isEmpty()) continue;

                double calories = getNumericValue(row.getCell(4));
                double proteins = getNumericValue(row.getCell(5));
                double fats     = getNumericValue(row.getCell(6));
                double carbs    = getNumericValue(row.getCell(7));
                double fiber    = getNumericValue(row.getCell(8));
                double sugars   = getNumericValue(row.getCell(9));

                foodDatabase.add(new Food(name, calories, proteins, fats, carbs, fiber, sugars));
            }
            return !foodDatabase.isEmpty();
        } catch (Exception e) {
            System.out.println("Error reading the Excel file: " + e.getMessage());
            return false;
        }
    }

    public void saveHistory() {
        try (PrintWriter out = new PrintWriter(new FileWriter(SAVE_FILE))) {
            out.println("GOALS|" + calorieGoal + "|" + proteinGoal + "|" + underMargin + "|" + overMargin);
            for (Map.Entry<String, List<MealEntry>> entry : historyByDate.entrySet()) {
                for (MealEntry m : entry.getValue()) {
                    out.println("MEAL|" + m.getDate() + "|" + m.getMealType() + "|" + m.getFood().getName() + "|" + m.getGrams());
                }
            }
        } catch (IOException ignored) {}
    }

    public void loadSavedHistory() {
        File f = new File(SAVE_FILE);
        if (!f.exists()) return;
        try (BufferedReader br = new BufferedReader(new FileReader(f))) {
            String line;
            while ((line = br.readLine()) != null) {
                String[] parts = line.split("\\|");
                if (parts.length == 5 && parts[0].equals("GOALS")) {
                    calorieGoal = Double.parseDouble(parts[1]);
                    proteinGoal = Double.parseDouble(parts[2]);
                    underMargin = Double.parseDouble(parts[3]);
                    overMargin = Double.parseDouble(parts[4]);
                } else if (parts.length == 5 && parts[0].equals("MEAL")) {
                    String date = parts[1];
                    String mealType = parts[2];
                    String foodName = parts[3];
                    double grams = Double.parseDouble(parts[4]);

                    Food matched = null;
                    for (Food dbFood : foodDatabase) {
                        if (dbFood.getName().equalsIgnoreCase(foodName)) {
                            matched = dbFood;
                            break;
                        }
                    }
                    if (matched != null) {
                        historyByDate.putIfAbsent(date, new ArrayList<>());
                        historyByDate.get(date).add(new MealEntry(date, mealType, matched, grams));
                    }
                }
            }
        } catch (Exception ignored) {}
    }

    private double getNumericValue(Cell cell) {
        if (cell == null) return 0.0;
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(cell.getStringCellValue().trim());
            } catch (NumberFormatException e) {
                return 0.0;
            }
        }
        return 0.0;
    }

    private String getStringValue(Cell cell) {
        if (cell == null) return "";
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim();
        }
        return "";
    }
}

// ============================================================================
