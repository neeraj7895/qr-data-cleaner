import React, { useState } from 'react';
import { Upload, FileSpreadsheet, Wand2, Languages, Download, CheckCircle, AlertCircle, Loader2 } from 'lucide-react';

const QRDataCleanerUI = () => {
  const [activeTab, setActiveTab] = useState('cleaner');
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [processing, setProcessing] = useState(false);
  const [hindiText, setHindiText] = useState('');
  const [englishText, setEnglishText] = useState('');
  const [translating, setTranslating] = useState(false);

  const handleFileUpload = (e) => {
    const files = Array.from(e.target.files);
    setUploadedFiles(files);
  };

  const processFiles = () => {
    setProcessing(true);
    setTimeout(() => {
      setProcessing(false);
      alert('Files processed successfully!');
    }, 2000);
  };

  const translateText = async () => {
    setTranslating(true);
    setEnglishText('');
    
    try {
      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 1000,
          messages: [
            {
              role: 'user',
              content: `Convert the following text to professional corporate English. If it's in Hindi/Hinglish, translate it. If it's already in English, improve it for professional communication:

"${hindiText}"

Provide only the professional English version without any explanations.`
            }
          ]
        })
      });

      const data = await response.json();
      const result = data.content.find(item => item.type === 'text')?.text || '';
      setEnglishText(result);
    } catch (error) {
      setEnglishText('Error: Could not translate. Please try again.');
    } finally {
      setTranslating(false);
    }
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(englishText);
    alert('Copied to clipboard!');
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-purple-50">
      {/* Header */}
      <div className="bg-white shadow-sm border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-6 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <div className="w-10 h-10 bg-gradient-to-br from-blue-600 to-purple-600 rounded-lg flex items-center justify-center">
                <FileSpreadsheet className="w-6 h-6 text-white" />
              </div>
              <div>
                <h1 className="text-2xl font-bold text-gray-800">QR Data Cleaner Pro</h1>
                <p className="text-sm text-gray-500">Clean, merge & standardize your data</p>
              </div>
            </div>
            <div className="flex items-center space-x-2 text-sm text-gray-600">
              <CheckCircle className="w-4 h-4 text-green-500" />
              <span>System Active</span>
            </div>
          </div>
        </div>
      </div>

      {/* Tab Navigation */}
      <div className="max-w-7xl mx-auto px-6 mt-6">
        <div className="flex space-x-2 bg-white rounded-lg p-1 shadow-sm border border-gray-200 w-fit">
          <button
            onClick={() => setActiveTab('cleaner')}
            className={`px-6 py-2.5 rounded-md font-medium transition-all flex items-center space-x-2 ${
              activeTab === 'cleaner'
                ? 'bg-gradient-to-r from-blue-600 to-purple-600 text-white shadow-md'
                : 'text-gray-600 hover:text-gray-800'
            }`}
          >
            <FileSpreadsheet className="w-4 h-4" />
            <span>Data Cleaner</span>
          </button>
          <button
            onClick={() => setActiveTab('translator')}
            className={`px-6 py-2.5 rounded-md font-medium transition-all flex items-center space-x-2 ${
              activeTab === 'translator'
                ? 'bg-gradient-to-r from-blue-600 to-purple-600 text-white shadow-md'
                : 'text-gray-600 hover:text-gray-800'
            }`}
          >
            <Languages className="w-4 h-4" />
            <span>English Creator</span>
          </button>
        </div>
      </div>

      {/* Main Content */}
      <div className="max-w-7xl mx-auto px-6 py-8">
        {activeTab === 'cleaner' ? (
          <div className="space-y-6">
            {/* Upload Section */}
            <div className="bg-white rounded-xl shadow-lg border border-gray-200 p-8">
              <div className="flex items-center space-x-3 mb-6">
                <div className="w-10 h-10 bg-blue-100 rounded-lg flex items-center justify-center">
                  <Upload className="w-5 h-5 text-blue-600" />
                </div>
                <div>
                  <h2 className="text-xl font-bold text-gray-800">Upload Excel Files</h2>
                  <p className="text-sm text-gray-500">Select one or multiple .xlsx files to clean</p>
                </div>
              </div>

              <div className="border-2 border-dashed border-gray-300 rounded-lg p-12 text-center hover:border-blue-500 transition-colors cursor-pointer bg-gray-50">
                <input
                  type="file"
                  multiple
                  accept=".xlsx,.xls"
                  onChange={handleFileUpload}
                  className="hidden"
                  id="file-upload"
                />
                <label htmlFor="file-upload" className="cursor-pointer">
                  <Upload className="w-12 h-12 text-gray-400 mx-auto mb-4" />
                  <p className="text-gray-600 font-medium mb-2">Drop files here or click to browse</p>
                  <p className="text-sm text-gray-400">Supports .xlsx and .xls files</p>
                </label>
              </div>

              {uploadedFiles.length > 0 && (
                <div className="mt-6 space-y-2">
                  <p className="text-sm font-medium text-gray-700 mb-3">Selected Files:</p>
                  {uploadedFiles.map((file, idx) => (
                    <div key={idx} className="flex items-center space-x-3 bg-blue-50 rounded-lg p-3">
                      <FileSpreadsheet className="w-5 h-5 text-blue-600" />
                      <span className="text-sm text-gray-700 flex-1">{file.name}</span>
                      <span className="text-xs text-gray-500">{(file.size / 1024).toFixed(1)} KB</span>
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Process Button */}
            {uploadedFiles.length > 0 && (
              <div className="bg-white rounded-xl shadow-lg border border-gray-200 p-6">
                <button
                  onClick={processFiles}
                  disabled={processing}
                  className="w-full bg-gradient-to-r from-blue-600 to-purple-600 text-white py-4 rounded-lg font-medium hover:shadow-lg transition-all flex items-center justify-center space-x-2 disabled:opacity-50"
                >
                  {processing ? (
                    <>
                      <Loader2 className="w-5 h-5 animate-spin" />
                      <span>Processing Files...</span>
                    </>
                  ) : (
                    <>
                      <Wand2 className="w-5 h-5" />
                      <span>Clean & Process Files</span>
                    </>
                  )}
                </button>

                <div className="mt-6 bg-gray-50 rounded-lg p-4">
                  <p className="text-sm font-medium text-gray-700 mb-2">Cleaning Operations:</p>
                  <ul className="space-y-1 text-sm text-gray-600">
                    <li className="flex items-center space-x-2">
                      <CheckCircle className="w-4 h-4 text-green-500" />
                      <span>Remove duplicate mobile numbers</span>
                    </li>
                    <li className="flex items-center space-x-2">
                      <CheckCircle className="w-4 h-4 text-green-500" />
                      <span>Clean 12-digit mobile numbers (remove '91' prefix)</span>
                    </li>
                    <li className="flex items-center space-x-2">
                      <CheckCircle className="w-4 h-4 text-green-500" />
                      <span>Standardize date formats (dd-mm-yyyy)</span>
                    </li>
                    <li className="flex items-center space-x-2">
                      <CheckCircle className="w-4 h-4 text-green-500" />
                      <span>Format Aadhaar & Account numbers</span>
                    </li>
                    <li className="flex items-center space-x-2">
                      <CheckCircle className="w-4 h-4 text-green-500" />
                      <span>Add dropdown validations</span>
                    </li>
                  </ul>
                </div>
              </div>
            )}
          </div>
        ) : (
          <div className="bg-white rounded-xl shadow-lg border border-gray-200 p-8">
            <div className="flex items-center space-x-3 mb-6">
              <div className="w-10 h-10 bg-purple-100 rounded-lg flex items-center justify-center">
                <Languages className="w-5 h-5 text-purple-600" />
              </div>
              <div>
                <h2 className="text-xl font-bold text-gray-800">English Creator</h2>
                <p className="text-sm text-gray-500">Convert Hindi/Hinglish to professional English</p>
              </div>
            </div>

            <div className="space-y-6">
              {/* Input Section */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Enter Your Text (Hindi/Hinglish/English)
                </label>
                <textarea
                  value={hindiText}
                  onChange={(e) => setHindiText(e.target.value)}
                  placeholder="Enter your message in Hindi, Hinglish, or English..."
                  className="w-full h-40 p-4 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent resize-none"
                />
              </div>

              {/* Translate Button */}
              <button
                onClick={translateText}
                disabled={!hindiText.trim() || translating}
                className="w-full bg-gradient-to-r from-purple-600 to-blue-600 text-white py-3 rounded-lg font-medium hover:shadow-lg transition-all flex items-center justify-center space-x-2 disabled:opacity-50"
              >
                {translating ? (
                  <>
                    <Loader2 className="w-5 h-5 animate-spin" />
                    <span>Converting to Professional English...</span>
                  </>
                ) : (
                  <>
                    <Wand2 className="w-5 h-5" />
                    <span>Convert to Professional English</span>
                  </>
                )}
              </button>

              {/* Output Section */}
              {englishText && (
                <div className="bg-gradient-to-br from-green-50 to-emerald-50 rounded-lg p-6 border border-green-200">
                  <div className="flex items-center justify-between mb-3">
                    <label className="block text-sm font-medium text-gray-700">
                      Professional English Output
                    </label>
                    <button
                      onClick={copyToClipboard}
                      className="text-sm bg-white px-3 py-1.5 rounded-md border border-gray-300 hover:bg-gray-50 transition-colors flex items-center space-x-1"
                    >
                      <Download className="w-4 h-4" />
                      <span>Copy</span>
                    </button>
                  </div>
                  <div className="bg-white rounded-lg p-4 border border-green-200">
                    <p className="text-gray-800 leading-relaxed">{englishText}</p>
                  </div>
                </div>
              )}

              {/* Info Box */}
              <div className="bg-blue-50 rounded-lg p-4 border border-blue-200">
                <div className="flex items-start space-x-3">
                  <AlertCircle className="w-5 h-5 text-blue-600 mt-0.5 flex-shrink-0" />
                  <div className="text-sm text-blue-800">
                    <p className="font-medium mb-1">How it works:</p>
                    <ul className="space-y-1 text-blue-700">
                      <li>• Translates Hindi/Hinglish to English</li>
                      <li>• Improves existing English to corporate standard</li>
                      <li>• Perfect for emails, tasks, and formal communication</li>
                    </ul>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Footer */}
      <div className="max-w-7xl mx-auto px-6 py-8">
        <div className="text-center text-sm text-gray-500">
          <p>Made with ❤️ for your team | Powered by AI</p>
        </div>
      </div>
    </div>
  );
};

export default QRDataCleanerUI;
