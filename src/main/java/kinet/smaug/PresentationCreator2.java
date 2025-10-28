package kinet.smaug;

import org.apache.poi.sl.usermodel.Insets2D;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.TextShape.TextAutofit;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.ThreadLocalRandom;
import java.util.regex.Pattern;

public final class PresentationCreator2 {

    // ========== КОНФИГУРАЦИЯ ==========
    private static final int SLIDE_W = 1920;
    private static final int SLIDE_H = 1080;

    private static final double MARGIN = 0;
    private static final double TITLE_H = 124;
    private static final double GAP = 40;

    private static final String FONT_FAMILY = "Segoe UI";
    private static final double TITLE_FONT_MAX = 56;
    private static final double TITLE_FONT_MIN = 26;
    private static final double BODY_FONT_MAX = 30;
    private static final double BODY_FONT_MIN = 15;

    private static final double BODY_LINE_H_K = 5.16;
    private static final double TITLE_LINE_H_K = 1.18;

    private static final double TEXT_WIDTH_RATIO_WITH_IMAGES = 0.40;
    private static final double TEXT_WIDTH_RATIO_NO_IMAGES = 0.92;

    // ========== ЦВЕТОВАЯ СХЕМА ==========
    private static final Color BG_PURPLE_DARK = new Color(0x1E, 0x14, 0x46);
    private static final Color BG_PURPLE_BAR = new Color(0x3A, 0x1F, 0x6F);
    private static final Color ACCENT_BAR = new Color(0xBB, 0x86, 0xFC);
    private static final Color TITLE_COLOR = new Color(0xF3, 0xE8, 0xFF);
    private static final Color BODY_COLOR = new Color(0xEE, 0xE9, 0xF7);
    private static final Color BULLET_COLOR = new Color(0xBB, 0x86, 0xFC);
    private static final Color FOOTER_COLOR = new Color(0xC8, 0xBE, 0xF0);

    // ========== НАСТРОЙКИ СЕТИ И КЭША ==========
    private static final Duration HTTP_CONNECT_TIMEOUT = Duration.ofSeconds(8);
    private static final Duration HTTP_REQUEST_TIMEOUT = Duration.ofSeconds(25);
    private static final int MAX_IMAGE_BYTES = 10 * 1024 * 1024;
    private static final Pattern URL_RE = Pattern.compile("^(?i)https?://.+");

    private static final String UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36";
    private static final String ACCEPT = "image/jpeg,image/png,image/gif,image/*;q=0.8,*/*;q=0.5";
    private static final String ACCEPT_LANG = "en-US,en;q=0.9";
    private static final String ACCEPT_ENC = "gzip, deflate, br";

    private static final Path CACHE_DIR = Paths.get("cache_images");


    // НАСТРОЙКА РЕЖИМА: при 1–2 картинках — кладём их ПОД текст, на всю ширину
    private static final boolean STACK_IMAGES_IF_LESS_OR_EQ_2 = true;
    // какая доля высоты контента отводится под текст в «стековом» режиме
    private static final double FEW_IMG_TEXT_HEIGHT_RATIO = 0.22;
    // внутренние поля, когда занимаем всю ширину
    private static final double CONTENT_SIDE_PADDING = 32;



    static {
        ImageIO.setUseCache(false);
        ImageIO.scanForPlugins();
        try {
            Files.createDirectories(CACHE_DIR);
        } catch (IOException e) {
            System.err.println("Не удалось создать каталог кэша: " + e.getMessage());
        }
    }

    // ========== ОСНОВНОЙ МЕТОД ==========
    public static void main(String[] args) {
        String inPath = (args != null && args.length >= 1) ? args[0] : "txt.txt";
        String outPath = (args != null && args.length >= 2) ? args[1] : "presentation.pptx";

        List<SlideSpec> slides;
        try {
            slides = parseSlides(new File(inPath));
        } catch (IOException e) {
            System.err.println("Не удалось прочитать входной файл: " + e.getMessage());
            return;
        }

        try (XMLSlideShow ppt = new XMLSlideShow()) {
            ppt.setPageSize(new Dimension(SLIDE_W, SLIDE_H));

            int page = 1;
            for (SlideSpec spec : slides) {
                XSLFSlide slide = ppt.createSlide();
                applyPurpleTheme(slide);

                // Расчет областей размещения контента
                double left = MARGIN;
                double top = MARGIN;
                double width = SLIDE_W - 2 * MARGIN;

                Rectangle2D titleBox = new Rectangle2D.Double(left, top, width, TITLE_H);
                addTitle(slide, spec.title, titleBox);

                double contentTop = top + TITLE_H + (GAP * 2);
                double contentHeight = SLIDE_H - contentTop - MARGIN;

                boolean hasImages = !spec.imageUrls.isEmpty();
                boolean hasText = !(spec.paragraphs.isEmpty() && spec.bullets.isEmpty());
                boolean fewImgs = hasImages && spec.imageUrls.size() <= 3 && STACK_IMAGES_IF_LESS_OR_EQ_2;

                Rectangle2D textBox;
                Rectangle2D imagesArea = null;

                if (fewImgs) {
                    // СТЕКОВЫЙ РЕЖИМ: текст сверху на всю ширину, картинки снизу по центру
                    double textH = contentHeight * FEW_IMG_TEXT_HEIGHT_RATIO;
                    double imgH  = contentHeight - textH;

                    // Текст — ширина почти во весь слайд
                    double textLeft = left + CONTENT_SIDE_PADDING;
                    double textWidth = width - 2 * CONTENT_SIDE_PADDING;

                    textBox = hasText
                            ? new Rectangle2D.Double(textLeft, contentTop, textWidth, textH - GAP * 0.5)
                            : new Rectangle2D.Double(0, 0, 0, 0);

                    // Картинки — на всю ширину слайдовой области, строго по центру
                    imagesArea = new Rectangle2D.Double(
                            left + CONTENT_SIDE_PADDING,
                            contentTop + textH + GAP * 0.5,
                            width - 2 * CONTENT_SIDE_PADDING,
                            imgH - GAP * 0.5
                    );

                    if (hasText) addBody(slide, spec, textBox);
                    if (hasImages) addImagesThreeColumns(slide, spec.imageUrls, imagesArea, ppt);
                } else {
                    // КОЛОНОЧНЫЙ РЕЖИМ: текст слева, картинки справа (для 3+ шт.)
                    double textAreaWidth = width * (hasImages ? TEXT_WIDTH_RATIO_WITH_IMAGES : TEXT_WIDTH_RATIO_NO_IMAGES);
                    double textAreaLeft = hasImages ? left : left + (width - textAreaWidth) / 2.0;

                    Rectangle2D tb = new Rectangle2D.Double(textAreaLeft, contentTop, textAreaWidth, contentHeight);
                    addBody(slide, spec, tb);

                    if (hasImages) {
                        double imgAreaLeft = left + textAreaWidth + GAP;
                        double imgAreaWidth = width - textAreaWidth - GAP;
                        imagesArea = new Rectangle2D.Double(imgAreaLeft, contentTop, imgAreaWidth, contentHeight);
                        addImagesThreeColumns(slide, spec.imageUrls, imagesArea, ppt);
                    }
                }


                addFooter(slide, page++, slides.size());
            }

            try (OutputStream os = new BufferedOutputStream(new FileOutputStream(outPath))) {
                ppt.write(os);
            }
            System.out.println("Готово: " + outPath);
        } catch (Exception e) {
            System.err.println("Ошибка при формировании презентации: " + e.getMessage());
            e.printStackTrace(System.err);
        }
    }

    // ========== ТЕМА И ОФОРМЛЕНИЕ ==========
    private static void applyPurpleTheme(XSLFSlide slide) {
        slide.getBackground().setFillColor(BG_PURPLE_DARK);

        XSLFAutoShape bar = slide.createAutoShape();
        bar.setShapeType(ShapeType.RECT);
        bar.setAnchor(new Rectangle((int) MARGIN, (int) MARGIN, (int) (SLIDE_W - 2 * MARGIN), (int) TITLE_H));
        bar.setFillColor(BG_PURPLE_BAR);
        bar.setLineWidth(0);

        XSLFAutoShape accent = slide.createAutoShape();
        accent.setShapeType(ShapeType.RECT);
        accent.setAnchor(new Rectangle((int) MARGIN, (int) (MARGIN + TITLE_H - 4), (int) (SLIDE_W - 2 * MARGIN), 4));
        accent.setFillColor(ACCENT_BAR);
        accent.setLineWidth(0);
    }

    private static void addFooter(XSLFSlide slide, int page, int total) {
        XSLFTextBox tb = slide.createTextBox();
        tb.setAnchor(new Rectangle((int) MARGIN, SLIDE_H - (int) (MARGIN * 0.6), SLIDE_W - (int) (2 * MARGIN), 22));
        tb.setTextAutofit(TextAutofit.NONE);
        tb.setVerticalAlignment(VerticalAlignment.MIDDLE);

        XSLFTextParagraph p = tb.addNewTextParagraph();
        p.setTextAlign(TextParagraph.TextAlign.RIGHT);
        XSLFTextRun r = p.addNewTextRun();
        r.setText(page + " / " + total);
        r.setFontFamily(FONT_FAMILY);
        r.setFontSize(11.0);
        r.setFontColor(FOOTER_COLOR);
    }

    // ========== ПАРСИНГ ВХОДНОГО ФАЙЛА ==========
    private static List<SlideSpec> parseSlides(File file) throws IOException {
        List<SlideSpec> out = new ArrayList<>();
        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(file), StandardCharsets.UTF_8))) {
            SlideSpec curr = null;
            String line;
            while ((line = br.readLine()) != null) {
                String s = line.strip();
                if (s.isEmpty()) continue;

                if (s.startsWith("# Слайд")) {
                    if (curr != null) out.add(curr);
                    String title = s;
                    int idx = s.indexOf(':');
                    if (idx >= 0 && idx + 1 < s.length()) title = s.substring(idx + 1).trim();
                    curr = new SlideSpec(title);
                } else if (s.startsWith("- ")) {
                    ensureSlide(curr);
                    curr.bullets.add(s.substring(2).trim());
                } else if (URL_RE.matcher(s).matches()) {
                    ensureSlide(curr);
                    String url = s;
                    if (url.contains("cloudinary.com") && url.contains("f_auto")) {
                        url = url.replace("f_auto", "f_jpg");
                    }
                    curr.imageUrls.add(url);
                } else {
                    ensureSlide(curr);
                    curr.paragraphs.add(s);
                }
            }
            if (curr != null) out.add(curr);
        }
        return out;
    }

    private static void ensureSlide(SlideSpec s) {
        if (s == null) throw new IllegalStateException("Файл начинается с контента без заголовка '# Слайд N: ...'");
    }

    // ========== ЗАГОЛОВОК СЛАЙДА ==========
    private static void addTitle(XSLFSlide slide, String title, Rectangle2D box) {
        XSLFTextBox tb = slide.createTextBox();
        tb.setAnchor(toRect(box));
        tb.clearText();
        tb.setTextAutofit(TextAutofit.NONE);
        tb.setWordWrap(true);
        tb.setVerticalAlignment(VerticalAlignment.MIDDLE);
        tb.setInsets(new Insets2D(8, 14, 8, 14));

        XSLFTextParagraph p = tb.addNewTextParagraph();
        p.setTextAlign(TextParagraph.TextAlign.LEFT);
        p.setSpaceBefore(0.0);
        p.setSpaceAfter(0.0);

        double fs = fitFontSize(title, box.getWidth() - 28, box.getHeight() - 16, TITLE_FONT_MAX, TITLE_FONT_MIN);
        List<String> lines = wrapSmart(title, charsPerLine(box.getWidth() - 28, fs));
        p.setLineSpacing(fs * TITLE_LINE_H_K);

        for (int i = 0; i < lines.size(); i++) {
            XSLFTextRun r = p.addNewTextRun();
            r.setText(lines.get(i));
            r.setFontSize(fs);
            r.setFontFamily(FONT_FAMILY);
            r.setFontColor(TITLE_COLOR);
            r.setBold(true);
            if (i < lines.size() - 1) p.addLineBreak();
        }
    }

    // ========== ОСНОВНОЙ ТЕКСТ СЛАЙДА ==========
    private static void addBody(XSLFSlide slide, SlideSpec spec, Rectangle2D box) {
        if (spec.paragraphs.isEmpty() && spec.bullets.isEmpty()) return;

        XSLFTextBox tb = slide.createTextBox();
        tb.setAnchor(toRect(box));
        tb.clearText();
        tb.setInsets(new Insets2D(20, 14, 10, 14));
        tb.setTextAutofit(TextAutofit.NONE);
        tb.setWordWrap(true);
        tb.setVerticalAlignment(VerticalAlignment.TOP);

        StringBuilder all = new StringBuilder();
        for (String s : spec.paragraphs) {
            if (!all.isEmpty()) all.append('\n');
            all.append(s);
        }
        for (String s : spec.bullets) {
            if (!all.isEmpty()) all.append('\n');
            all.append(s);
        }

        double fs = fitFontSize(all.toString(), box.getWidth() - 28, box.getHeight() - 20, BODY_FONT_MAX, BODY_FONT_MIN);
        int cpl = charsPerLine(box.getWidth() - 28, fs);

        // Абзацы
        for (String para : spec.paragraphs) {
            List<String> lines = wrapSmart(para, cpl);
            XSLFTextParagraph p = tb.addNewTextParagraph();
            p.setTextAlign(TextParagraph.TextAlign.LEFT);
            p.setBullet(false);
            p.setSpaceBefore(0.0);
            p.setSpaceAfter(5.0);
            p.setLineSpacing(fs * BODY_LINE_H_K);

            for (int i = 0; i < lines.size(); i++) {
                XSLFTextRun r = p.addNewTextRun();
                r.setText(lines.get(i));
                r.setFontSize(fs);
                r.setFontFamily(FONT_FAMILY);
                r.setFontColor(BODY_COLOR);
                if (i < lines.size() - 1) p.addLineBreak();
            }
        }

        // Маркированные списки
        for (String bullet : spec.bullets) {
            List<String> lines = wrapSmart(bullet, Math.max(12, cpl - 2));
            XSLFTextParagraph p = tb.addNewTextParagraph();
            p.setTextAlign(TextParagraph.TextAlign.LEFT);
            p.setBullet(true);
            p.setBulletCharacter("•");
            p.setBulletFontColor(BULLET_COLOR);
            p.setLeftMargin(24.0);
            p.setIndent(-12.0);
            p.setSpaceBefore(0.0);
            p.setSpaceAfter(0.0);
            p.setLineSpacing(fs * BODY_LINE_H_K);

            for (int i = 0; i < lines.size(); i++) {
                XSLFTextRun r = p.addNewTextRun();
                r.setText(lines.get(i));
                r.setFontSize(fs);
                r.setFontFamily(FONT_FAMILY);
                r.setFontColor(BODY_COLOR);
                if (i < lines.size() - 1) p.addLineBreak();
            }
        }
    }

    // ========== КАРТИНКИ (3 КОЛОНКИ С ЦЕНТРИРОВАНИЕМ) ==========
// ========== КАРТИНКИ (ЦЕНТРИРОВАННЫЕ) ==========
    private static void addImagesThreeColumns(XSLFSlide slide, List<String> urls, Rectangle2D area, XMLSlideShow ppt) {
        int count = Math.min(urls.size(), 3);
        if (count <= 0) return;

        // Размеры и отступы
        double imgGap = 20; // Отступ между картинками
        double availableHeight = area.getHeight() * 0.45; // Используем 85% высоты
        double imgHeight = availableHeight;
        double imgWidth = imgHeight * 16.0 / 9.0; // Соотношение 16:9

        // Общая ширина всех картинок с отступами
        double totalWidth = (imgWidth * count) + (imgGap * (count - 1));

        // Стартовая позиция X для центрирования всей группы
        double startX = area.getX() + (area.getWidth() - totalWidth) / 2;

        // Позиция Y для вертикального центрирования
        double startY = area.getY() + (area.getHeight() - imgHeight) / 2;

        for (int i = 0; i < count; i++) {
            double x = startX + i * (imgWidth + imgGap);
            Rectangle2D imgRect = new Rectangle2D.Double(x, startY, imgWidth, imgHeight);

            String url = urls.get(i);
            try {
                BufferedImage img = downloadImageWithCache(url);
                if (img == null) {
                    System.err.println("Пропуск: не удалось получить " + url);
                    continue;
                }

                BufferedImage cropped = cropToAspect(img, 16.0 / 9.0);

                // Тень
                Rectangle2D shadow = new Rectangle2D.Double(
                        imgRect.getX() + 6,
                        imgRect.getY() + 6,
                        imgRect.getWidth() - 12,
                        imgRect.getHeight() - 12
                );
                XSLFAutoShape sh = slide.createAutoShape();
                sh.setShapeType(ShapeType.ROUND_RECT);
                sh.setAnchor(toRect(shadow));
                sh.setFillColor(new Color(0, 0, 0, 50));
                sh.setLineWidth(0);

                // Картинка
                XSLFPictureData pd = ppt.addPicture(toPngBytes(cropped), PictureData.PictureType.PNG);
                XSLFPictureShape pic = slide.createPicture(pd);
                pic.setAnchor(toRect(imgRect));
                pic.setLineColor(new Color(255, 255, 255, 60));
                pic.setLineWidth(1.5);

            } catch (Exception e) {
                System.err.println("Не удалось добавить изображение: " + url + " (" + e.getMessage() + ")");
            }
        }
    }

    private static BufferedImage cropToAspect(BufferedImage src, double aspect) {
        int w = src.getWidth(), h = src.getHeight();
        double cur = (double) w / h;
        int x = 0, y = 0, cw = w, ch = h;

        if (cur > aspect) {
            cw = (int) Math.round(h * aspect);
            x = (w - cw) / 2;
        } else if (cur < aspect) {
            ch = (int) Math.round(w / aspect);
            y = (h - ch) / 2;
        }

        BufferedImage sub = src.getSubimage(x, y, Math.max(1, cw), Math.max(1, ch));
        BufferedImage out = new BufferedImage(sub.getWidth(), sub.getHeight(), BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = out.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g.drawImage(sub, 0, 0, null);
        g.dispose();
        return out;
    }

    private static byte[] toPngBytes(BufferedImage img) throws IOException {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            ImageIO.write(img, "png", baos);
            return baos.toByteArray();
        }
    }

    // ========== РАБОТА С ИЗОБРАЖЕНИЯМИ И КЭШИРОВАНИЕ ==========
    private static BufferedImage downloadImageWithCache(String url) throws IOException, InterruptedException {
        if (!URL_RE.matcher(url).matches()) return null;
        byte[] bytes = getBytesCached(url);
        if (bytes == null) return null;
        if (bytes.length == 0 || bytes.length > MAX_IMAGE_BYTES) return null;

        try (ByteArrayInputStream bis = new ByteArrayInputStream(bytes)) {
            BufferedImage img = ImageIO.read(bis);
            if (img != null) return img;
        }

        ImageIO.scanForPlugins();
        try (ByteArrayInputStream bis = new ByteArrayInputStream(bytes)) {
            return ImageIO.read(bis);
        }
    }

    private static byte[] getBytesCached(String url) throws IOException, InterruptedException {
        String key = sha256Hex(url);
        Path target = CACHE_DIR.resolve(key + ".bin");

        if (Files.isReadable(target)) {
            try {
                byte[] data = Files.readAllBytes(target);
                if (data.length > 0) return data;
            } catch (IOException e) {
                System.err.println("Ошибка чтения кэша: " + target + " — " + e.getMessage());
            }
        }

        byte[] fresh = fetchBytesWithRetries(url);
        if (fresh == null) return null;

        Path tmp = CACHE_DIR.resolve(key + ".tmp");
        try (OutputStream os = Files.newOutputStream(tmp, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING, StandardOpenOption.WRITE)) {
            os.write(fresh);
        }

        try {
            Files.move(tmp, target, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
        } catch (AtomicMoveNotSupportedException ex) {
            Files.move(tmp, target, StandardCopyOption.REPLACE_EXISTING);
        }

        return fresh;
    }

    private static String sha256Hex(String s) {
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            byte[] dig = md.digest(s.getBytes(StandardCharsets.UTF_8));
            StringBuilder sb = new StringBuilder(dig.length * 2);
            for (byte b : dig) sb.append(String.format("%02x", b));
            return sb.toString();
        } catch (NoSuchAlgorithmException e) {
            throw new RuntimeException(e);
        }
    }

    private static byte[] fetchBytesWithRetries(String url) throws IOException, InterruptedException {
        java.net.URI uri = java.net.URI.create(url);
        String host = uri.getHost() == null ? "" : uri.getHost().toLowerCase();
        String referer = host.endsWith("gstatic.com") || host.contains("google") ?
                "https://www.google.com/" : "https://" + host + "/";

        IOException lastIoEx = null;
        int lastSc = 0;

        for (int attempt = 1; attempt <= 3; attempt++) {
            try {
                HttpURLConnection conn = (HttpURLConnection) new URL(url).openConnection();
                conn.setRequestMethod("GET");
                conn.setConnectTimeout((int) HTTP_CONNECT_TIMEOUT.toMillis());
                conn.setReadTimeout((int) HTTP_REQUEST_TIMEOUT.toMillis());
                conn.setInstanceFollowRedirects(true);
                conn.setRequestProperty("User-Agent", UA);
                conn.setRequestProperty("Accept", ACCEPT);
                conn.setRequestProperty("Accept-Language", ACCEPT_LANG);
                conn.setRequestProperty("Accept-Encoding", ACCEPT_ENC);
                conn.setRequestProperty("Referer", referer);

                int sc = conn.getResponseCode();
                lastSc = sc;

                if (sc >= 200 && sc < 300) {
                    String ct = conn.getContentType();
                    if (ct != null && ct.toLowerCase().contains("text/html")) {
                        System.err.println("Пропуск (HTML вместо изображения): " + url + " [CT=" + ct + "]");
                        return null;
                    }

                    try (InputStream is = conn.getInputStream();
                         ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
                        byte[] buffer = new byte[8192];
                        int len, total = 0;
                        while ((len = is.read(buffer)) != -1) {
                            baos.write(buffer, 0, len);
                            total += len;
                            if (total > MAX_IMAGE_BYTES) {
                                System.err.println("Изображение слишком большое: " + url);
                                return null;
                            }
                        }
                        byte[] body = baos.toByteArray();
                        return body.length > 0 ? body : null;
                    }
                } else if (sc == 401 || sc == 403 || sc == 404) {
                    System.err.println("Доступ закрыт/не найден: " + url + " [HTTP " + sc + "]");
                    return null;
                } else if (sc == 429 || (sc >= 500 && sc < 600)) {
                    backoff(attempt);
                } else {
                    System.err.println("Статус " + sc + " для: " + url);
                    return null;
                }
            } catch (IOException io) {
                lastIoEx = io;
                System.err.println("IOException на попытке " + attempt + ": " + io.getMessage());
                backoff(attempt);
            }
        }

        if (lastIoEx != null) throw lastIoEx;
        System.err.println("Не удалось получить: " + url + " [HTTP " + lastSc + "]");
        return null;
    }

    private static void backoff(int attempt) {
        try {
            long base = 250L * attempt;
            long jitter = ThreadLocalRandom.current().nextLong(0, 200);
            Thread.sleep(base + jitter);
        } catch (InterruptedException ignored) {}
    }

    // ========== РАБОТА С ТЕКСТОМ И ШРИФТАМИ ==========
    private static double fitFontSize(String text, double boxW, double boxH, double maxPt, double minPt) {
        double lo = minPt, hi = maxPt, best = minPt;
        for (int iter = 0; iter < 18; iter++) {
            double mid = (lo + hi) * 0.5;
            if (fits(text, boxW, boxH, mid)) {
                best = mid;
                lo = mid;
            } else {
                hi = mid;
            }
            if (Math.abs(hi - lo) < 0.25) break;
        }
        return clamp(best, minPt, maxPt);
    }

    private static boolean fits(String text, double boxW, double boxH, double pt) {
        int cpl = charsPerLine(boxW, pt);
        List<String> lines = new ArrayList<>();
        for (String para : text.split("\n")) lines.addAll(wrapSmart(para, cpl));
        double lineH = pt * BODY_LINE_H_K;
        double totalH = lines.size() * lineH;
        return totalH <= boxH && maxLineLength(lines) <= cpl;
    }

    private static int charsPerLine(double boxW, double pt) {
        double avg = pt * 0.52;
        int cpl = (int) Math.floor((boxW - 10) / Math.max(5.5, avg));
        return Math.max(12, cpl);
    }

    private static int maxLineLength(List<String> lines) {
        int m = 0;
        for (String l : lines) m = Math.max(m, l.length());
        return m;
    }

    private static List<String> wrapSmart(String text, int cpl) {
        List<String> out = new ArrayList<>();
        if (text == null || text.isBlank()) {
            out.add("");
            return out;
        }

        String[] words = text.trim().split("\\s+");
        StringBuilder line = new StringBuilder();

        for (String w : words) {
            if (line.length() == 0) {
                appendFitting(out, line, w, cpl);
            } else if (line.length() + 1 + w.length() <= cpl) {
                line.append(' ').append(w);
            } else {
                out.add(line.toString());
                line.setLength(0);
                appendFitting(out, line, w, cpl);
            }
        }
        if (line.length() > 0) out.add(line.toString());

        // Анти-сирота
        if (out.size() >= 2) {
            int lastLen = out.get(out.size() - 1).length();
            if (lastLen < Math.max(6, cpl * 0.25)) {
                String prev = out.get(out.size() - 2);
                int cut = prev.lastIndexOf(' ');
                if (cut > 0) {
                    String move = prev.substring(cut + 1);
                    out.set(out.size() - 2, prev.substring(0, cut));
                    out.set(out.size() - 1, move + " " + out.get(out.size() - 1));
                }
            }
        }
        return out;
    }

    private static void appendFitting(List<String> out, StringBuilder line, String word, int cpl) {
        if (word.length() <= cpl) {
            line.append(word);
            return;
        }

        int cut = Math.max(Math.max(word.lastIndexOf('-'), word.lastIndexOf('/')), word.lastIndexOf('.'));
        if (cut > 0 && cut < word.length() - 1) {
            String left = word.substring(0, cut + 1);
            String right = word.substring(cut + 1);
            if (left.length() > cpl) {
                forceHyphenate(out, line, left, cpl);
            } else {
                line.append(left);
                out.add(line.toString());
                line.setLength(0);
                appendFitting(out, line, right, cpl);
            }
            return;
        }
        forceHyphenate(out, line, word, cpl);
    }

    private static void forceHyphenate(List<String> out, StringBuilder line, String word, int cpl) {
        int idx = Math.min(cpl - 1, Math.max(3, word.length() / 2));
        String left = word.substring(0, idx) + "-";
        String right = word.substring(idx);
        if (line.length() > 0) {
            out.add(line.toString());
            line.setLength(0);
        }
        out.add(left);
        appendFitting(out, line, right, cpl);
    }

    // ========== УТИЛИТЫ ==========
    private static Rectangle toRect(Rectangle2D r2d) {
        return new Rectangle(
                (int) Math.round(r2d.getX()),
                (int) Math.round(r2d.getY()),
                (int) Math.round(r2d.getWidth()),
                (int) Math.round(r2d.getHeight())
        );
    }

    private static Rectangle2D fitRect(Dimension natural, Rectangle2D box) {
        double iw = natural.getWidth(), ih = natural.getHeight();
        if (iw <= 0 || ih <= 0) return new Rectangle2D.Double(box.getX(), box.getY(), 1, 1);

        double scale = Math.min(box.getWidth() / iw, box.getHeight() / ih);
        double w = Math.max(1, iw * scale), h = Math.max(1, ih * scale);
        double x = box.getX() + (box.getWidth() - w) / 2.0;
        double y = box.getY() + (box.getHeight() - h) / 2.0;

        return new Rectangle2D.Double(x, y, w, h);
    }

    private static double clamp(double v, double lo, double hi) {
        return Math.max(lo, Math.min(hi, v));
    }

    // ========== МОДЕЛЬ ДАННЫХ ==========
    private static final class SlideSpec {
        final String title;
        final List<String> paragraphs = new ArrayList<>();
        final List<String> bullets = new ArrayList<>();
        final List<String> imageUrls = new ArrayList<>();

        SlideSpec(String title) {
            this.title = Objects.requireNonNullElse(title, "");
        }
    }
}