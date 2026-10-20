import 'dart:convert';
import 'dart:typed_data';

import 'package:dio/dio.dart';
import 'package:file_picker/file_picker.dart';
import 'package:file_saver/file_saver.dart';
import 'package:flutter/material.dart';

void main() => runApp(const RaporExcelApp());

class RaporExcelApp extends StatelessWidget {
  const RaporExcelApp({super.key});

  @override
  Widget build(BuildContext context) {
    const navy = Color(0xFF152A46);
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Rapor → Excel',
      theme: ThemeData(
        useMaterial3: true,
        colorScheme: ColorScheme.fromSeed(seedColor: const Color(0xFF246BFD)),
        scaffoldBackgroundColor: const Color(0xFFF4F7FB),
        appBarTheme: const AppBarTheme(
          backgroundColor: navy,
          foregroundColor: Colors.white,
          centerTitle: false,
        ),
        inputDecorationTheme: InputDecorationTheme(
          filled: true,
          fillColor: Colors.white,
          border: OutlineInputBorder(borderRadius: BorderRadius.circular(14)),
          enabledBorder: OutlineInputBorder(
            borderRadius: BorderRadius.circular(14),
            borderSide: const BorderSide(color: Color(0xFFDCE4EF)),
          ),
        ),
      ),
      home: const ReportHomePage(),
    );
  }
}

class ReportHomePage extends StatefulWidget {
  const ReportHomePage({super.key});

  @override
  State<ReportHomePage> createState() => _ReportHomePageState();
}

class _ReportHomePageState extends State<ReportHomePage> {
  // IMPORTANT: Set this to the HTTPS address of your own backend API.
  // The backend must accept multipart fields "pdf" and "reportType" and
  // return the completed .xlsm file as the response body.
  static const String apiUrl = 'https://SIZIN-SERVISINIZ.onrender.com/generate-excel';
  // Prototip tokeni. Gerçek kullanımda APK içine gömülen token güvenli kabul edilmez.
  static const String apiToken = 'BURAYA_APP_API_TOKEN';

  final Dio _dio = Dio(
    BaseOptions(
      connectTimeout: const Duration(seconds: 30),
      receiveTimeout: const Duration(minutes: 5),
      sendTimeout: const Duration(minutes: 5),
    ),
  );

  PlatformFile? _selectedPdf;
  String _reportType = 'auto';
  bool _isProcessing = false;
  double _progress = 0;
  String? _statusMessage;
  bool _statusIsError = false;

  Future<void> _pickPdf() async {
    try {
      final result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: const ['pdf'],
        withData: false,
      );
      if (result == null || result.files.isEmpty) return;
      final file = result.files.single;
      if (file.path == null) {
        _showStatus('PDF dosyasına erişilemedi. Dosyayı tekrar seçin.', error: true);
        return;
      }
      setState(() {
        _selectedPdf = file;
        _statusMessage = null;
        _progress = 0;
      });
    } catch (e) {
      _showStatus('Dosya seçilemedi: $e', error: true);
    }
  }

  Future<void> _generateExcel() async {
    final pdf = _selectedPdf;
    if (pdf == null || pdf.path == null) {
      _showStatus('Önce bir PDF raporu seçin.', error: true);
      return;
    }
    if (apiUrl.contains('SIZIN-SERVISINIZ') || apiToken == 'BURAYA_APP_API_TOKEN') {
      _showStatus(
        'Sunucu adresi veya uygulama tokenı ayarlanmamış. main.dart içindeki apiUrl ve apiToken alanlarını düzenleyin.',
        error: true,
      );
      return;
    }

    setState(() {
      _isProcessing = true;
      _progress = 0;
      _statusMessage = 'PDF gönderiliyor ve rapor analiz ediliyor…';
      _statusIsError = false;
    });

    try {
      final formData = FormData.fromMap({
        'pdf': await MultipartFile.fromFile(
          pdf.path!,
          filename: pdf.name,
          contentType: DioMediaType('application', 'pdf'),
        ),
        // auto = server detects MMO, Royal Cert or another supported layout.
        'reportType': _reportType,
      });

      final response = await _dio.post<List<int>>(
        apiUrl,
        data: formData,
        options: Options(
          responseType: ResponseType.bytes,
          headers: {
            'Accept': 'application/vnd.ms-excel, application/octet-stream',
            'X-App-Token': apiToken,
          },
          validateStatus: (status) => status != null && status >= 200 && status < 600,
        ),
        onSendProgress: (sent, total) {
          if (total > 0 && mounted) {
            setState(() => _progress = sent / total * 0.45);
          }
        },
        onReceiveProgress: (received, total) {
          if (total > 0 && mounted) {
            setState(() => _progress = 0.45 + received / total * 0.5);
          }
        },
      );

      final bytes = Uint8List.fromList(response.data ?? <int>[]);
      if (response.statusCode == null || response.statusCode! < 200 || response.statusCode! >= 300) {
        String detail = 'Sunucu Excel dosyasını oluşturamadı (HTTP ${response.statusCode}).';
        try {
          final decoded = jsonDecode(utf8.decode(bytes));
          if (decoded is Map && decoded['detail'] != null) detail = decoded['detail'].toString();
          if (decoded is Map && decoded['error'] != null) detail = decoded['error'].toString();
        } catch (_) {}
        throw Exception(detail);
      }
      if (bytes.isEmpty) throw Exception('Sunucu boş dosya döndürdü.');

      setState(() {
        _progress = 0.98;
        _statusMessage = 'Excel dosyası telefona kaydediliyor…';
      });

      final baseName = pdf.name.replaceAll(RegExp(r'\.pdf$', caseSensitive: false), '');
      final safeName = '${baseName}_doldurulmus';
      await FileSaver.instance.saveFile(
        name: safeName,
        bytes: bytes,
        ext: 'xlsm',
        mimeType: MimeType.microsoftExcel,
      );

      if (!mounted) return;
      setState(() {
        _progress = 1;
        _statusMessage = 'Tamamlandı! Doldurulmuş Excel dosyası kaydedildi.';
        _statusIsError = false;
      });
    } on DioException catch (e) {
      final message = e.response?.statusMessage ?? e.message ?? 'Sunucuya bağlanılamadı.';
      _showStatus('Bağlantı hatası: $message', error: true);
    } catch (e) {
      _showStatus(e.toString().replaceFirst('Exception: ', ''), error: true);
    } finally {
      if (mounted) setState(() => _isProcessing = false);
    }
  }

  void _showStatus(String message, {required bool error}) {
    if (!mounted) return;
    setState(() {
      _statusMessage = message;
      _statusIsError = error;
    });
  }

  String _sizeLabel(int? bytes) {
    if (bytes == null) return 'PDF dosyası';
    if (bytes < 1024 * 1024) return '${(bytes / 1024).toStringAsFixed(0)} KB';
    return '${(bytes / (1024 * 1024)).toStringAsFixed(1)} MB';
  }

  @override
  Widget build(BuildContext context) {
    const navy = Color(0xFF152A46);
    return Scaffold(
      appBar: AppBar(
        title: const Text('Rapor → Otomatik Excel', style: TextStyle(fontWeight: FontWeight.w700)),
        actions: [
          IconButton(
            tooltip: 'Nasıl çalışır?',
            onPressed: _showHelp,
            icon: const Icon(Icons.help_outline_rounded),
          ),
        ],
      ),
      body: SafeArea(
        child: ListView(
          padding: const EdgeInsets.fromLTRB(18, 20, 18, 28),
          children: [
            Container(
              padding: const EdgeInsets.all(22),
              decoration: BoxDecoration(
                gradient: const LinearGradient(
                  colors: [Color(0xFF1D3D66), Color(0xFF246BFD)],
                  begin: Alignment.topLeft,
                  end: Alignment.bottomRight,
                ),
                borderRadius: BorderRadius.circular(24),
              ),
              child: const Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Icon(Icons.auto_awesome_rounded, color: Colors.white, size: 34),
                  SizedBox(height: 14),
                  Text('PDF raporunu Excel’e dönüştür',
                      style: TextStyle(color: Colors.white, fontSize: 23, height: 1.15, fontWeight: FontWeight.w800)),
                  SizedBox(height: 8),
                  Text('MMO, Royal Cert ve desteklenen rapor formatlarını analiz edip Excel şablonunu otomatik doldur.',
                      style: TextStyle(color: Colors.white, height: 1.4, fontSize: 14)),
                ],
              ),
            ),
            const SizedBox(height: 22),
            const Text('1. Rapor türü', style: TextStyle(fontSize: 16, fontWeight: FontWeight.w800, color: navy)),
            const SizedBox(height: 9),
            DropdownButtonFormField<String>(
              value: _reportType,
              decoration: const InputDecoration(prefixIcon: Icon(Icons.document_scanner_outlined)),
              items: const [
                DropdownMenuItem(value: 'auto', child: Text('Otomatik algıla (MMO / Royal Cert)')),
                DropdownMenuItem(value: 'mmo', child: Text('MMO raporu')),
                DropdownMenuItem(value: 'royal_cert', child: Text('Royal Cert raporu')),
              ],
              onChanged: _isProcessing ? null : (value) => setState(() => _reportType = value ?? 'auto'),
            ),
            const SizedBox(height: 20),
            const Text('2. PDF raporu', style: TextStyle(fontSize: 16, fontWeight: FontWeight.w800, color: navy)),
            const SizedBox(height: 9),
            InkWell(
              borderRadius: BorderRadius.circular(20),
              onTap: _isProcessing ? null : _pickPdf,
              child: Container(
                padding: const EdgeInsets.all(18),
                decoration: BoxDecoration(
                  color: Colors.white,
                  borderRadius: BorderRadius.circular(20),
                  border: Border.all(color: _selectedPdf == null ? const Color(0xFFD6E0EF) : const Color(0xFF246BFD), width: 1.4),
                ),
                child: Row(
                  children: [
                    Container(
                      height: 54,
                      width: 54,
                      decoration: BoxDecoration(color: const Color(0xFFFFEEEE), borderRadius: BorderRadius.circular(15)),
                      child: const Icon(Icons.picture_as_pdf_rounded, color: Color(0xFFD92D20), size: 29),
                    ),
                    const SizedBox(width: 13),
                    Expanded(
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Text(_selectedPdf?.name ?? 'PDF dosyası seç', maxLines: 2, overflow: TextOverflow.ellipsis,
                              style: const TextStyle(fontWeight: FontWeight.w700, fontSize: 15)),
                          const SizedBox(height: 4),
                          Text(_selectedPdf == null ? 'Telefonundan MMO veya Royal Cert raporu seç' : _sizeLabel(_selectedPdf!.size),
                              style: const TextStyle(color: Color(0xFF68778D), fontSize: 12)),
                        ],
                      ),
                    ),
                    const Icon(Icons.folder_open_rounded, color: Color(0xFF246BFD)),
                  ],
                ),
              ),
            ),
            const SizedBox(height: 18),
            const _InfoRow(icon: Icons.manage_search_rounded, title: 'Raporu analiz et', detail: 'Bina bilgileri, tarihler, standartlar ve uygunsuzluk kodları okunur.'),
            const _InfoRow(icon: Icons.fact_check_outlined, title: 'Proforma kodlarını eşleştir', detail: 'Şablondaki kod listesine göre açıklamalar bulunur; eşleşmeyenler işaretlenir.'),
            const _InfoRow(icon: Icons.file_download_outlined, title: 'Excel’i kaydet', detail: 'Doldurulmuş .xlsm dosyası telefona indirilir.'),
            const SizedBox(height: 20),
            if (_isProcessing) ...[
              ClipRRect(
                borderRadius: BorderRadius.circular(20),
                child: LinearProgressIndicator(value: _progress.clamp(0.0, 1.0), minHeight: 7),
              ),
              const SizedBox(height: 9),
              Text(_statusMessage ?? 'İşleniyor…', textAlign: TextAlign.center,
                  style: const TextStyle(color: Color(0xFF53647A), fontSize: 13)),
              const SizedBox(height: 14),
            ],
            SizedBox(
              height: 56,
              child: FilledButton.icon(
                onPressed: _isProcessing ? null : _generateExcel,
                icon: _isProcessing
                    ? const SizedBox(width: 20, height: 20, child: CircularProgressIndicator(strokeWidth: 2, color: Colors.white))
                    : const Icon(Icons.auto_awesome_rounded),
                label: Text(_isProcessing ? 'Rapor işleniyor…' : 'Excel’i otomatik hazırla',
                    style: const TextStyle(fontSize: 16, fontWeight: FontWeight.w700)),
                style: FilledButton.styleFrom(shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16))),
              ),
            ),
            if (_statusMessage != null && !_isProcessing) ...[
              const SizedBox(height: 14),
              Container(
                padding: const EdgeInsets.all(14),
                decoration: BoxDecoration(
                  color: _statusIsError ? const Color(0xFFFFF0EF) : const Color(0xFFEAF8EF),
                  borderRadius: BorderRadius.circular(15),
                  border: Border.all(color: _statusIsError ? const Color(0xFFF3C1BC) : const Color(0xFFB7E3C5)),
                ),
                child: Row(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Icon(_statusIsError ? Icons.error_outline_rounded : Icons.check_circle_outline_rounded,
                        color: _statusIsError ? const Color(0xFFB42318) : const Color(0xFF18794E)),
                    const SizedBox(width: 10),
                    Expanded(child: Text(_statusMessage!, style: TextStyle(color: _statusIsError ? const Color(0xFF912018) : const Color(0xFF17663F)))),
                  ],
                ),
              ),
            ],
            const SizedBox(height: 18),
            const Text(
              'Önemli: PDF’yi okuyup Excel’i gerçekten dolduran yapay zekâ/işlem motoru sunucu tarafında çalışır. Bu uygulama dosyayı güvenli HTTPS API’nize gönderir; API adresi ve sunucu kodu kurulmadan otomatik dönüştürme çalışmaz.',
              style: TextStyle(fontSize: 12, color: Color(0xFF6C7A8D), height: 1.45),
            ),
          ],
        ),
      ),
    );
  }

  void _showHelp() {
    showDialog<void>(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text('Nasıl çalışır?'),
        content: const Text(
          '1. Rapor türünü seçin veya otomatik algılamayı kullanın.\n\n'
          '2. MMO ya da Royal Cert PDF dosyasını seçin.\n\n'
          '3. “Excel’i otomatik hazırla” düğmesine basın.\n\n'
          '4. Sunucu raporu analiz eder, proforma listesiyle eşleştirir ve doldurulmuş Excel dosyasını geri gönderir.\n\n'
          'İlk kurulumda apiUrl adresi ve rapor işleme sunucusu ayarlanmalıdır.',
        ),
        actions: [TextButton(onPressed: () => Navigator.pop(context), child: const Text('Tamam'))],
      ),
    );
  }
}

class _InfoRow extends StatelessWidget {
  const _InfoRow({required this.icon, required this.title, required this.detail});
  final IconData icon;
  final String title;
  final String detail;

  @override
  Widget build(BuildContext context) {
    return Padding(
      padding: const EdgeInsets.only(bottom: 13),
      child: Row(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Container(
            width: 40,
            height: 40,
            decoration: BoxDecoration(color: const Color(0xFFE7EEFF), borderRadius: BorderRadius.circular(12)),
            child: Icon(icon, color: const Color(0xFF246BFD), size: 22),
          ),
          const SizedBox(width: 12),
          Expanded(
            child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
              Text(title, style: const TextStyle(fontWeight: FontWeight.w700, fontSize: 14)),
              const SizedBox(height: 3),
              Text(detail, style: const TextStyle(color: Color(0xFF68778D), fontSize: 12, height: 1.35)),
            ]),
          ),
        ],
      ),
    );
  }
}
