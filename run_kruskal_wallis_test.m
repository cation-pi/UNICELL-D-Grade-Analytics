% This script performs the Kruskal-Wallis test on the aggregate data to determine if there are any statistically significant differences among the medians of the various ADCA product grades. 
% It iterates through each physicochemical parameter and exports a summary report with the test statistics and the hypothesis decision ("Reject" or "Fail to Reject") for each.

inputFile = 'masterData_KW.xlsx';
outputFile = 'Laporan_Uji_Kruskal_Wallis.xlsx';

if exist(outputFile, 'file')
    delete(outputFile);
    fprintf('Info: File laporan lama "%s" telah dihapus untuk memastikan hasil baru.\n', outputFile);
end

try
    dataTable = readtable(inputFile);
    fprintf('Berhasil memuat file data: %s\n', inputFile);
catch ME
    errorMessage = sprintf(['Error: File "%s" tidak ditemukan.\n' ...
        'Pastikan file tersebut berada di folder yang sama dengan skrip ini.\n' ...
        'Pesan Error MATLAB: %s'], inputFile, ME.message);
    errordlg(errorMessage, 'File Tidak Ditemukan');
    return;
end

dependentVars = {'FI_DT_mean', 'FI_GV_mean', 'FI_APS_mean', 'FI_MC_mean', 'FI_PH_mean'};
alpha = 0.05;

summaryData = table('Size', [numel(dependentVars), 5], ...
                    'VariableTypes', {'string', 'double', 'double', 'double', 'string'}, ...
                    'VariableNames', {'Parameter', 'Chi_Square', 'df', 'p_value', 'Keputusan_H0'});

fprintf('Struktur laporan Excel telah disiapkan. Memulai analisis...\n');
fprintf('------------------------------------------------------------\n\n');

for i = 1:length(dependentVars)
    currentVar = dependentVars{i};
    fprintf('Menganalisis parameter: %s...\n', currentVar);
    
    [p, tbl, stats] = kruskalwallis(dataTable.(currentVar), dataTable.GRADE, 'off');
    
    chi_square_val = tbl{2, 5};
    df_val = tbl{2, 3};
    p_val = tbl{2, 6};
    
    if p_val < alpha
        keputusan = "Ditolak";
    else
        keputusan = "Gagal Ditolak";
    end
    
    summaryData(i, :) = {currentVar, chi_square_val, df_val, p_val, keputusan};
    
    detailSheetName = ['Detail_' currentVar(4:end-5)];
    
    detailTable = cell2table(tbl(2:end,:), 'VariableNames', tbl(1,:));
    
    writetable(detailTable, outputFile, 'Sheet', detailSheetName);
end

fprintf('\n------------------------------------------------------------\n');
fprintf('Semua parameter telah dianalisis.\nMenyusun file laporan akhir...\n');

writetable(summaryData, outputFile, 'Sheet', 'Ringkasan Hasil');

infoSheetData = {
    'Judul Laporan', 'Laporan Hasil Uji Statistik Kruskal-Wallis';
    'File Data Sumber', inputFile;
    'Tanggal Analisis', datestr(now, 'dd-mmm-yyyy HH:MM:SS');
    'Tingkat Signifikansi (α)', alpha
};
writecell(infoSheetData, outputFile, 'Sheet', 'Informasi Analisis', 'Range', 'A1');

fprintf('Laporan lengkap telah berhasil dibuat dan disimpan sebagai: %s\n', outputFile);
