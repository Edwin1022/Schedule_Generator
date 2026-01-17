import React, { useState } from 'react';
import { 
  StyleSheet, Text, View, TextInput, TouchableOpacity, 
  ActivityIndicator, Alert, SafeAreaView, Platform, 
  KeyboardAvoidingView, ScrollView, TouchableWithoutFeedback, Keyboard 
} from 'react-native';
import * as DocumentPicker from 'expo-document-picker';
import * as FileSystem from 'expo-file-system/legacy';
import * as Sharing from 'expo-sharing';

// ⚠️ REPLACE '172.20.10.2' with your actual IP if it changes!
const BASE_URL = 'http://172.20.10.2:5001'; 

export default function HomeScreen() {
  const [file, setFile] = useState<DocumentPicker.DocumentPickerAsset | null>(null);
  const [viewType, setViewType] = useState<'room' | 'teacher'>('room');
  const [outputName, setOutputName] = useState('My_Schedule');
  const [loading, setLoading] = useState(false);

  const pickDocument = async () => {
    try {
      const result = await DocumentPicker.getDocumentAsync({
        type: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'],
        copyToCacheDirectory: true,
      });

      if (!result.canceled && result.assets && result.assets.length > 0) {
        setFile(result.assets[0]);
      }
    } catch (err) {
      Alert.alert("Error", "Failed to pick file");
    }
  };

  const handleGenerate = async () => {
    if (!file) {
      Alert.alert("Missing File", "Please select an Excel file first.");
      return;
    }

    setLoading(true);
    
    const endpoint = viewType === 'room' ? '/api/schedule/room' : '/api/schedule/teacher';
    const apiUrl = `${BASE_URL}${endpoint}`;
    const filename = outputName.endsWith('.xlsx') ? outputName : `${outputName}.xlsx`;

    try {
      const formData = new FormData();
      // @ts-ignore
      formData.append('file', {
        uri: file.uri,
        name: file.name,
        type: file.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      formData.append('filename', filename);

      console.log(`🚀 Uploading to: ${apiUrl}`);

      const response = await fetch(apiUrl, {
        method: 'POST',
        body: formData,
        headers: { 'Content-Type': 'multipart/form-data' },
      });

      console.log("📡 Response Status:", response.status);

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Server Error ${response.status}: ${errorText}`);
      }

      const blob = await response.blob();
      const reader = new FileReader();
      reader.readAsDataURL(blob);
      reader.onloadend = async () => {
        const base64data = (reader.result as string).split(',')[1];
        // Ensure documentDirectory is not null by defaulting to ''
        const fileUri = (FileSystem.documentDirectory || '') + filename;

        await FileSystem.writeAsStringAsync(fileUri, base64data, {
          encoding: 'base64',
        });

        setLoading(false);

        if (await Sharing.isAvailableAsync()) {
          await Sharing.shareAsync(fileUri);
        } else {
          Alert.alert("Success", `File saved to: ${fileUri}`);
        }
      };

    } catch (error: any) {
      console.error(error);
      setLoading(false);
      Alert.alert("Error", error.message || "Something went wrong");
    }
  };

  return (
    <SafeAreaView style={styles.safeArea}>
      {/* KeyboardAvoidingView pushes UI up when keyboard opens */}
      <KeyboardAvoidingView 
        behavior={Platform.OS === 'ios' ? 'padding' : 'height'}
        style={styles.keyboardContainer}
      >
        {/* TouchableWithoutFeedback dismisses keyboard when tapping outside */}
        <TouchableWithoutFeedback onPress={Keyboard.dismiss}>
          <ScrollView 
            contentContainerStyle={styles.scrollContainer} 
            keyboardShouldPersistTaps="handled"
          >
            
            <Text style={styles.header}>📅 Schedule Generator</Text>

            {/* --- 1. VIEW TOGGLE --- */}
            <Text style={styles.label}>Select Output View:</Text>
            <View style={styles.toggleContainer}>
              <TouchableOpacity 
                style={[styles.toggleBtn, viewType === 'room' && styles.toggleBtnActive]}
                onPress={() => setViewType('room')}
              >
                <Text style={[styles.toggleText, viewType === 'room' && styles.activeText]}>Room View</Text>
              </TouchableOpacity>
              <TouchableOpacity 
                style={[styles.toggleBtn, viewType === 'teacher' && styles.toggleBtnActive]}
                onPress={() => setViewType('teacher')}
              >
                <Text style={[styles.toggleText, viewType === 'teacher' && styles.activeText]}>Teacher View</Text>
              </TouchableOpacity>
            </View>

            {/* --- 2. FILE PICKER --- */}
            <Text style={styles.label}>Input File:</Text>
            <TouchableOpacity style={styles.uploadBox} onPress={pickDocument}>
              <Text style={styles.uploadText}>
                {file ? `📄 ${file.name}` : "📂 Tap to select Excel file"}
              </Text>
            </TouchableOpacity>

            {/* --- 3. FILENAME INPUT --- */}
            <Text style={styles.label}>Output Filename:</Text>
            <TextInput 
              style={styles.input}
              value={outputName}
              onChangeText={setOutputName}
              placeholder="e.g. Schedule_Final"
              placeholderTextColor="#A0AEC0"
            />

            {/* --- 4. SUBMIT BUTTON --- */}
            <TouchableOpacity 
              style={[styles.generateBtn, loading && styles.disabledBtn]} 
              onPress={handleGenerate}
              disabled={loading}
            >
              {loading ? (
                <ActivityIndicator color="#fff" />
              ) : (
                <Text style={styles.generateBtnText}>🚀 Generate & Download</Text>
              )}
            </TouchableOpacity>

          </ScrollView>
        </TouchableWithoutFeedback>
      </KeyboardAvoidingView>
    </SafeAreaView>
  );
}

const styles = StyleSheet.create({
  safeArea: {
    flex: 1,
    backgroundColor: '#F5F7FA',
  },
  keyboardContainer: {
    flex: 1,
  },
  scrollContainer: {
    padding: 24,
    paddingTop: 40,
    justifyContent: 'center',
    minHeight: '100%', // Ensures content is centered if small
  },
  header: {
    fontSize: 28,
    fontWeight: 'bold',
    color: '#2D3748',
    marginBottom: 32,
    textAlign: 'center',
  },
  label: {
    fontSize: 16,
    fontWeight: '600',
    color: '#4A5568',
    marginBottom: 8,
    marginTop: 16,
  },
  toggleContainer: {
    flexDirection: 'row',
    backgroundColor: '#E2E8F0',
    borderRadius: 8,
    padding: 4,
  },
  toggleBtn: {
    flex: 1,
    paddingVertical: 10,
    alignItems: 'center',
    borderRadius: 6,
  },
  toggleBtnActive: {
    backgroundColor: '#fff',
    shadowColor: '#000',
    shadowOffset: { width: 0, height: 1 },
    shadowOpacity: 0.1,
    shadowRadius: 2,
    elevation: 2,
  },
  toggleText: {
    color: '#718096',
    fontWeight: '600',
  },
  activeText: {
    color: '#2D3748',
    fontWeight: 'bold',
  },
  uploadBox: {
    borderWidth: 2,
    borderColor: '#CBD5E0',
    borderStyle: 'dashed',
    borderRadius: 12,
    padding: 24,
    alignItems: 'center',
    backgroundColor: '#fff',
  },
  uploadText: {
    color: '#4A5568',
    fontSize: 15,
  },
  input: {
    backgroundColor: '#fff',
    borderWidth: 1,
    borderColor: '#E2E8F0',
    borderRadius: 8,
    padding: 12,
    fontSize: 16,
    color: '#2D3748',
  },
  generateBtn: {
    backgroundColor: '#4299E1',
    paddingVertical: 16,
    borderRadius: 8,
    alignItems: 'center',
    marginTop: 32,
    shadowColor: '#4299E1',
    shadowOffset: { width: 0, height: 4 },
    shadowOpacity: 0.3,
    shadowRadius: 4,
    elevation: 4,
    marginBottom: 40, // Extra space at bottom for scrolling
  },
  disabledBtn: {
    backgroundColor: '#A0AEC0',
  },
  generateBtnText: {
    color: '#fff',
    fontSize: 18,
    fontWeight: 'bold',
  },
});