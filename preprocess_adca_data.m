% This script handles the initial data preparation. The main process involves cleaning the dataset by verifying completeness, handling missing values, and standardizing data formats to produce the final master datasets
% Final master dataset: masterData_KW.xlsx (aggregate) and masterData_FM.xlsx (non-aggregate).

namaFile_KW = 'Dataset_KW_notclean.xlsx';
namaFile_FM = 'Dataset_Sekunder_FM_notclean.xlsx';
namaOutput_KW = 'masterData_KW.xlsx';
namaOutput_FM = 'masterData_FM.xlsx';

try
    data_KW = readtable(namaFile_KW);
    fprintf('Berhasil memuat file: %s\n', namaFile_KW);
    
    data_FM = readtable(namaFile_FM);
    fprintf('Berhasil memuat file: %s\n', namaFile_FM);
catch ME
    errorMessage = sprintf(['Error: Tidak dapat memuat file Excel.\n' ...
        'Pastikan file "%s" dan "%s" ada di folder yang sama dengan skrip ini.\n' ...
        'Pesan error MATLAB: %s'], namaFile_KW, namaFile_FM, ME.message);
    uiwait(warndlg(errorMessage));
    return;
end

fprintf('\n--- Memproses Data Kruskal-Wallis ---\n');
[masterData_KW, ringkasan_KW] = cleanAndStandardizeData(data_KW);
fprintf('Ukuran data awal: %d baris, %d kolom.\n', ringkasan_KW.UkuranAwal(1), ringkasan_KW.UkuranAwal(2));
fprintf('Ukuran data setelah pembersihan: %d baris, %d kolom.\n', ringkasan_KW.UkuranAkhir(1), ringkasan_KW.UkuranAkhir(2));
fprintf('%d baris telah dihapus.\n', ringkasan_KW.BarisDihapus);

fprintf('\n--- Memproses Data Friedman ---\n');
[masterData_FM, ringkasan_FM] = cleanAndStandardizeData(data_FM);
fprintf('Ukuran data awal: %d baris, %d kolom.\n', ringkasan_FM.UkuranAwal(1), ringkasan_FM.UkuranAwal(2));
fprintf('Ukuran data setelah pembersihan: %d baris, %d kolom.\n', ringkasan_FM.UkuranAkhir(1), ringkasan_FM.UkuranAkhir(2));
fprintf('%d baris telah dihapus.\n', ringkasan_FM.BarisDihapus);

try
    writetable(masterData_KW, namaOutput_KW);
    fprintf('\nBerhasil menyimpan data bersih ke: %s\n', namaOutput_KW);
    
    writetable(masterData_FM, namaOutput_FM);
    fprintf('Berhasil menyimpan data bersih ke: %s\n', namaOutput_FM);
catch ME
    errorMessage = sprintf(['Error: Tidak dapat menyimpan file hasil.\n' ...
        'Pastikan Anda memiliki izin untuk menulis file di folder ini.\n' ...
        'Pesan error MATLAB: %s'], ME.message);
    uiwait(warndlg(errorMessage));
    return;
end

fprintf('\nSkrip pembersihan data telah selesai dijalankan.\n');

function [cleanedTable, summary] = cleanAndStandardizeData(inputTable)
    summary.UkuranAwal = size(inputTable);
    
    vars = inputTable.Properties.VariableNames;
    for i = 1:length(vars)
        if isnumeric(inputTable.(vars{i}))
            zero_indices = (inputTable.(vars{i}) == 0);
            inputTable.(vars{i})(zero_indices) = NaN;
        end
    end
    
    tempTable = rmmissing(inputTable);
    
    if ismember('GRADE', tempTable.Properties.VariableNames)
        if ~iscell(tempTable.GRADE) && ~isstring(tempTable.GRADE)
             tempTable.GRADE = cellstr(num2str(tempTable.GRADE));
        end
        tempTable.GRADE = upper(tempTable.GRADE);
        tempTable.GRADE = erase(tempTable.GRADE, ' ');
        fprintf('Kolom "GRADE" telah distandarkan.\n');
    else
        fprintf('Info: Kolom "GRADE" tidak ditemukan, langkah standardisasi dilewati.\n');
    end
    
    if ismember('DATE', tempTable.Properties.VariableNames)
        if ~isdatetime(tempTable.DATE)
            try
                tempTable.DATE = datetime(tempTable.DATE, 'InputFormat', 'dd/MM/yyyy');
                fprintf('Kolom "DATE" telah diformat sebagai datetime.\n');
            catch
                fprintf('Peringatan: Gagal mengonversi kolom "DATE". Periksa formatnya.\n');
            end
        end
        tempTable.DATE.Format = 'dd/MM/yyyy';
    else
        fprintf('Info: Kolom "DATE" tidak ditemukan, langkah format tanggal dilewati.\n');
    end
    
    cleanedTable = tempTable;
    
    summary.UkuranAkhir = size(cleanedTable);
    summary.BarisDihapus = summary.UkuranAwal(1) - summary.UkuranAkhir(1);
end
