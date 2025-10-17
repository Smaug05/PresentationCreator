package kinet.smaug;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * Проходит по всем подпапкам в data/ и создаёт out/<safeName>.pptx
 * - 20 слайдов из text.txt (первый блок = первый слайд)
 * - последний слайд — «Источники» (URL из текста + sources.txt)
 * - картинки: !image: <file> в блоке ИЛИ перемешанный список картинок из папки
 */
public final class PresentationCreator {

    // === Конфиг ===
    private static final String DATA_DIR = "data";
    private static final String DEFAULT_TEXT_FILE = "text.txt";
    private static final String OPTIONAL_SOURCES_FILE = "sources.txt";
    private static final String OPTIONAL_TEMPLATE = "templates/modern.pptx"; // если файл есть — используем, иначе чистый pptx

    private static final int REQUIRED_BODY_SLIDES = 20; // ровно 20
    private static final Dimension SLIDE_SIZE = new Dimension(1280, 720);

    // стиль
    private static final String FONT_TITLE = "Calibri";
    private static final String FONT_BODY  = "Calibri";
    private static final double TITLE_SIZE = 44;
    private static final double BODY_SIZE  = 24;

    // акцентная полоса слева
    private static final Color ACCENT    = new Color(75, 56, 137);
    private static final Color GRAY_TEXT = new Color(33, 37, 41);

    // разметка
    private static final double MARGIN       = 36;
    private static final double LEFT_COL_RATIO = 0.52; // текст ~52% ширины
    private static final double TITLE_TOP    = 24;

    private PresentationCreator() {}

    public static void main(String[] args) {
        new PresentationCreator().run();
    }

    private void run() {
        final Path dataRoot = Paths.get(DATA_DIR).toAbsolutePath().normalize();
        ensureDir(dataRoot, "data root");

        List<Path> dirs;
        try {
            dirs = Files.list(dataRoot)
                    .filter(Files::isDirectory)
                    .sorted(Comparator.comparing(p -> p.getFileName().toString().toLowerCase(Locale.ROOT)))
                    .collect(Collectors.toList());
        } catch (IOException e) {
            fatal("Cannot list data directory: " + e.getMessage(), e);
            return;
        }

        if (dirs.isEmpty()) {
            fatal("Нет подпапок в " + dataRoot + ". Создай хотя бы data/<имя>/ с файлами text.txt и, по желанию, sources.txt и картинками.");
            return;
        }

        System.out.println("Найдено папок: " + dirs.size());
        for (Path dir : dirs) {
            buildOne(dir);
        }
    }

    private void buildOne(Path presFolder) {
        final String rawName = presFolder.getFileName().toString();
        final String safeName = sanitizeName(rawName);
        if (safeName.isEmpty()) {
            System.err.println("Пропуск: некорректное имя папки: " + rawName);
            return;
        }

        final Path outPath = Paths.get("out", safeName + ".pptx");
        ensureDir(outPath.getParent(), "output dir");

        if (!presFolder.startsWith(Paths.get(DATA_DIR).toAbsolutePath().normalize())) {
            System.err.println("Пропуск: вне data/: " + presFolder);
            return;
        }

        // --- читаем текст ---
        final Path textFile = presFolder.resolve(DEFAULT_TEXT_FILE);
        if (!Files.isRegularFile(textFile)) {
            System.err.println("Пропуск: нет text.txt в " + presFolder);
            return;
        }
        final List<String> rawBlocks = readBlocks(textFile);
        final List<SlideBlock> slides = normalize(parseBlocks(rawBlocks), REQUIRED_BODY_SLIDES);

        // --- источники ---
        final Set<String> sources = collectSources(slides);
        final Path sourcesFile = presFolder.resolve(OPTIONAL_SOURCES_FILE);
        if (Files.isRegularFile(sourcesFile)) {
            try {
                Files.lines(sourcesFile, StandardCharsets.UTF_8)
                        .map(String::trim).filter(s -> !s.isEmpty())
                        .forEach(sources::add);
            } catch (IOException e) {
                System.err.println("Warning: can't read sources.txt: " + e.getMessage());
            }
        }

        // --- картинки ---
        List<Path> images = listImages(presFolder);
        // перемешиваем на каждом запуске (новый порядок), но стабилизируем в пределах папки через hash имени
        if (!images.isEmpty()) {
            long seed = System.currentTimeMillis() + safeName.hashCode();
            Collections.shuffle(images, new Random(seed));
            System.out.println("[" + safeName + "] images shuffled (" + images.size() + ")");
        }
        final List<Path> fallbackImages = images; // может быть пустым

        // --- генерация PPTX ---
        try (XMLSlideShow ppt = loadTemplateOrBlank(OPTIONAL_TEMPLATE)) {
            ppt.setPageSize(SLIDE_SIZE);

            for (int i = 0; i < REQUIRED_BODY_SLIDES; i++) {
                SlideBlock b = slides.get(i);
                Path explicit = resolveExplicitImage(b, presFolder);
                Path img = (explicit != null) ? explicit
                        : (!fallbackImages.isEmpty() ? fallbackImages.get(i % fallbackImages.size()) : null);
                createContentSlide(ppt, b, img, i + 1); // первый блок = первый слайд
            }

            createSourcesSlide(ppt, sources);

            Files.createDirectories(outPath.getParent());
            try (var os = Files.newOutputStream(outPath, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
                ppt.write(os);
            }
            System.out.println("OK: " + outPath.toAbsolutePath());
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("Generate error (" + safeName + "): " + e.getMessage());
        }
    }

    // ---------- Модель блока ----------
    private static final Pattern IMAGE_DIRECTIVE = Pattern.compile("^!image\\s*:\\s*(.+)$", Pattern.CASE_INSENSITIVE);

    private static class SlideBlock {
        String title;                 // заголовок (первая строка, # снимается)
        List<String> bullets = new ArrayList<>();  // остальные строки (буллеты)
        String raw;                   // сырой блок на всякий
        String explicitImage;         // из !image: file.jpg
    }

    private static List<String> readBlocks(Path textFile) {
        try {
            String txt = Files.readString(textFile, StandardCharsets.UTF_8);
            return Arrays.stream(txt.split("\\s*//\\s*"))
                    .map(String::trim).filter(s -> !s.isEmpty()).collect(Collectors.toList());
        } catch (IOException e) {
            throw new RuntimeException("Read text error: " + e.getMessage(), e);
        }
    }

    private static List<SlideBlock> parseBlocks(List<String> blocks) {
        List<SlideBlock> out = new ArrayList<>();
        for (String block : blocks) {
            SlideBlock b = new SlideBlock();
            b.raw = block;

            List<String> lines = Arrays.stream(block.split("\\R"))
                    .map(String::trim).filter(s -> !s.isEmpty()).collect(Collectors.toList());

            // image директива
            lines.removeIf(line -> {
                Matcher m = IMAGE_DIRECTIVE.matcher(line);
                if (m.find()) { b.explicitImage = m.group(1).trim(); return true; }
                return false;
            });

            if (lines.isEmpty()) {
                b.title = "Слайд";
                out.add(b);
                continue;
            }

            String first = lines.get(0);
            if (first.startsWith("#")) b.title = first.replaceFirst("^#+\\s*", "").trim();
            else b.title = first;

            for (int i = 1; i < lines.size(); i++) b.bullets.add(lines.get(i));

            out.add(b);
        }
        return out;
    }

    private static List<SlideBlock> normalize(List<SlideBlock> raw, int target) {
        List<SlideBlock> res = new ArrayList<>(target);
        if (raw.isEmpty()) {
            for (int i = 0; i < target; i++) {
                SlideBlock b = new SlideBlock();
                b.title = "Слайд " + (i + 1);
                b.bullets.add("Содержимое слайда " + (i + 1));
                res.add(b);
            }
        } else if (raw.size() >= target) {
            res.addAll(raw.subList(0, target));
        } else {
            res.addAll(raw);
            for (int i = raw.size(); i < target; i++) {
                SlideBlock b = new SlideBlock();
                b.title = "Слайд " + (i + 1);
                b.bullets.add("Содержимое слайда " + (i + 1));
                res.add(b);
            }
        }
        return res;
    }

    private static Set<String> collectSources(List<SlideBlock> slides) {
        Pattern url = Pattern.compile("(https?://\\S+)");
        Set<String> set = new LinkedHashSet<>();
        for (SlideBlock b : slides) {
            Matcher m1 = url.matcher(b.title);
            while (m1.find()) set.add(trimPunct(m1.group(1)));
            for (String s : b.bullets) {
                Matcher m2 = url.matcher(s);
                while (m2.find()) set.add(trimPunct(m2.group(1)));
            }
        }
        return set;
    }

    private static String trimPunct(String u) {
        return u.replaceAll("[)\\]\\},.;!?]+$", "");
    }

    private static Path resolveExplicitImage(SlideBlock b, Path presFolder) {
        if (b.explicitImage == null) return null;
        Path p = presFolder.resolve(b.explicitImage).normalize();
        return Files.isRegularFile(p) ? p : null;
    }

    // ---------- PPTX ----------

    private XMLSlideShow loadTemplateOrBlank(String templatePath) {
        Path p = Paths.get(templatePath);
        if (Files.isRegularFile(p)) {
            try (InputStream is = Files.newInputStream(p)) {
                return new XMLSlideShow(is);
            } catch (IOException e) {
                System.err.println("Warning: can't load template, fallback to blank: " + e.getMessage());
            }
        }
        return new XMLSlideShow();
    }

    private void createContentSlide(XMLSlideShow ppt, SlideBlock b, Path imagePath, int index) throws IOException {
        XSLFSlideMaster master = ppt.getSlideMasters().get(0);
        XSLFSlideLayout cl = getLayout(master, SlideLayout.TITLE_AND_CONTENT);
        XSLFSlide slide = ppt.createSlide(cl);

        addAccentStripe(slide);

        // заголовок
        XSLFTextShape title = slide.getPlaceholder(0);
        if (title != null) {
            title.clearText();
            styleTitleBox(title);
            addRun(title, b.title, FONT_TITLE, 34, true, GRAY_TEXT);
        }

        // тело
        XSLFTextShape content = slide.getPlaceholder(1);
        if (content != null) {
            content.clearText();
            styleBodyBox(content);
            for (String line : b.bullets) {
                if (line.isBlank()) continue;
                XSLFTextParagraph p = content.addNewTextParagraph();
                p.setBullet(true);
                p.setLeftMargin(28.0);
                p.setIndent(-14.0);
                p.setSpaceAfter(4.0);
                p.setFontAlign(TextParagraph.FontAlign.AUTO);

                XSLFTextRun run = p.addNewTextRun();
                run.setText(line);
                run.setFontFamily(FONT_BODY);
                run.setFontSize(BODY_SIZE);
                run.setFontColor(GRAY_TEXT);

                addHyperlinksIfAny(p);
            }
        }

        // картинка
        if (imagePath != null && Files.isRegularFile(imagePath)) {
            insertImage(ppt, slide, imagePath);
        }

        // номер слайда
        addFooter(slide, index);
    }

    private void createSourcesSlide(XMLSlideShow ppt, Set<String> sources) {
        XSLFSlideMaster master = ppt.getSlideMasters().get(0);
        XSLFSlideLayout cl = getLayout(master, SlideLayout.TITLE_AND_CONTENT);
        XSLFSlide slide = ppt.createSlide(cl);

        addAccentStripe(slide);

        XSLFTextShape title = slide.getPlaceholder(0);
        if (title != null) {
            title.clearText();
            styleTitleBox(title);
            addRun(title, "Источники", FONT_TITLE, 34, true, GRAY_TEXT);
        }

        XSLFTextShape content = slide.getPlaceholder(1);
        if (content != null) {
            content.clearText();
            styleBodyBox(content);
            if (sources.isEmpty()) {
                XSLFTextParagraph p = content.addNewTextParagraph();
                XSLFTextRun r = p.addNewTextRun();
                r.setText("Источники не указаны.");
                r.setFontFamily(FONT_BODY);
                r.setFontSize(BODY_SIZE);
                r.setFontColor(GRAY_TEXT);
            } else {
                for (String src : sources) {
                    XSLFTextParagraph p = content.addNewTextParagraph();
                    p.setBullet(true);
                    p.setLeftMargin(28.0);
                    p.setIndent(-14.0);
                    XSLFTextRun r = p.addNewTextRun();
                    r.setText(src);
                    r.setFontFamily(FONT_BODY);
                    r.setFontSize(18.0);
                    r.setFontColor(GRAY_TEXT);
                    try {
                        XSLFHyperlink link = r.createHyperlink();
                        link.setAddress(src);
                    } catch (Exception ignore) {}
                }
            }
        }
    }

    private XSLFSlideLayout getLayout(XSLFSlideMaster master, SlideLayout type) {
        try {
            return master.getLayout(type);
        } catch (Exception e) {
            return master.getSlideLayouts()[0];
        }
    }

    private void addAccentStripe(XSLFSlide slide) {
        XSLFAutoShape stripe = slide.createAutoShape();
        stripe.setShapeType(ShapeType.RECT);
        stripe.setFillColor(ACCENT);
        stripe.setLineColor(ACCENT);
        stripe.setAnchor(new Rectangle2D.Double(0, 0, 10, SLIDE_SIZE.getHeight()));
    }

    private void styleTitleBox(XSLFTextShape t) {
        t.setVerticalAlignment(VerticalAlignment.TOP);
        t.setAnchor(new Rectangle2D.Double(MARGIN, TITLE_TOP, SLIDE_SIZE.getWidth() * (LEFT_COL_RATIO - 0.06), 90));
        t.clearText();
    }

    private void styleBodyBox(XSLFTextShape t) {
        double x = MARGIN;
        double y = 120;
        double w = SLIDE_SIZE.getWidth() * LEFT_COL_RATIO - MARGIN;
        double h = SLIDE_SIZE.getHeight() - y - MARGIN;
        t.setAnchor(new Rectangle2D.Double(x, y, w, h));
        t.clearText();
    }

    private void addRun(XSLFTextShape box, String text, String font, double size, boolean bold, Color color) {
        XSLFTextParagraph p = box.addNewTextParagraph();
        XSLFTextRun r = p.addNewTextRun();
        r.setText(text);
        r.setFontFamily(font);
        r.setFontSize(size);
        r.setBold(bold);
        r.setFontColor(color);
    }

    private void addHyperlinksIfAny(XSLFTextParagraph p) {
        Pattern url = Pattern.compile("(https?://\\S+)");
        for (XSLFTextRun r : p.getTextRuns()) {
            Matcher m = url.matcher(r.getRawText());
            if (m.find()) {
                try {
                    XSLFHyperlink link = r.createHyperlink();
                    link.setAddress(trimPunct(m.group(1)));
                } catch (Exception ignore) {}
            }
        }
    }

    private void insertImage(XMLSlideShow ppt, XSLFSlide slide, Path imagePath) throws IOException {
        byte[] bytes = Files.readAllBytes(imagePath);
        PictureData.PictureType type = detectPictureType(imagePath);
        XSLFPictureData picData = ppt.addPicture(bytes, type);
        XSLFPictureShape pic = slide.createPicture(picData);

        double left = SLIDE_SIZE.getWidth() * LEFT_COL_RATIO + MARGIN;
        double top = 96;
        double maxW = SLIDE_SIZE.getWidth() - left - MARGIN;
        double maxH = SLIDE_SIZE.getHeight() - top - MARGIN;

        BufferedImage bi = ImageIO.read(imagePath.toFile());
        if (bi != null) {
            double iw = bi.getWidth();
            double ih = bi.getHeight();
            double scale = Math.min(maxW / iw, maxH / ih);
            double w = Math.floor(iw * scale);
            double h = Math.floor(ih * scale);
            pic.setAnchor(new Rectangle2D.Double(left, top, w, h));
        } else {
            pic.setAnchor(new Rectangle2D.Double(left, top, maxW, maxH));
        }
    }

    private void addFooter(XSLFSlide slide, int index) {
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(SLIDE_SIZE.getWidth() - 80, SLIDE_SIZE.getHeight() - 28, 72, 20));
        XSLFTextRun r = box.addNewTextParagraph().addNewTextRun();
        r.setText(String.valueOf(index));
        r.setFontFamily(FONT_BODY);
        r.setFontSize(12.0);
        r.setFontColor(new Color(100, 106, 115));
    }

    private static PictureData.PictureType detectPictureType(Path p) {
        String n = p.getFileName().toString().toLowerCase(Locale.ROOT);
        if (n.endsWith(".png")) return PictureData.PictureType.PNG;
        if (n.endsWith(".jpg") || n.endsWith(".jpeg")) return PictureData.PictureType.JPEG;
        if (n.endsWith(".gif")) return PictureData.PictureType.GIF;
        return PictureData.PictureType.PNG;
    }

    // ---------- FS/утилиты ----------

    private static void ensureDir(Path p, String what) {
        try { if (p != null) Files.createDirectories(p); }
        catch (IOException e) { throw new RuntimeException("Cannot create " + what + ": " + e.getMessage(), e); }
    }

    private static String sanitizeName(String raw) {
        String s = raw.replaceAll("[^\\p{IsAlphabetic}\\p{IsDigit}._-]+", "");
        if (s.equals(".") || s.equals("..")) return "";
        return s;
    }

    private static List<Path> listImages(Path folder) {
        try {
            return Files.list(folder)
                    .filter(Files::isRegularFile)
                    .filter(p -> {
                        String n = p.getFileName().toString().toLowerCase(Locale.ROOT);
                        return n.endsWith(".jpg") || n.endsWith(".jpeg") || n.endsWith(".png") || n.endsWith(".gif");
                    })
                    .sorted(Comparator.comparing(PresentationCreator::naturalKeyString))
                    .collect(Collectors.toList());
        } catch (IOException e) {
            return List.of();
        }
    }

    // Естественная сортировка: img2 < img10 (используется для стабильности до перемешивания)
    private static String naturalKeyString(Path p) {
        String s = p.getFileName().toString().toLowerCase(Locale.ROOT);
        StringBuilder sb = new StringBuilder();
        Matcher m = Pattern.compile("\\d+|\\D+").matcher(s);
        while (m.find()) {
            String part = m.group();
            if (Character.isDigit(part.charAt(0))) {
                String num = part.replaceFirst("^0+(?!$)", "");
                sb.append(String.format("%020d", new java.math.BigInteger(num)));
            } else {
                sb.append(part);
            }
        }
        return sb.toString();
    }

    // --- fatal helpers ---
    private static void fatal(String msg) {
        System.err.println(msg);
        System.exit(1);
    }
    private static void fatal(String msg, Throwable ex) {
        if (ex != null) ex.printStackTrace();
        System.err.println(msg);
        System.exit(1);
    }
}
