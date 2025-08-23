% Performs descriptive statistical analysis and visualization. It generates comparative box plots from non-aggregate data to visualize variability and internal consistency. 
% Concurrently, it calculates descriptive statistics (mean, median, IQR) from aggregate data.
% It also generates a ranked report of the grades based on their consistency metrics.

fileFM = 'masterData_FM.xlsx'; 
fileKW = 'masterData_KW.xlsx'; 
folderPlots = 'Plots';
folderResults = 'Results';
fileOutputStats = 'StatDeskriptif_Agregat.xlsx';

if ~exist(folderPlots, 'dir'), mkdir(folderPlots); fprintf('Folder "%s" telah dibuat.\n', folderPlots); end
if ~exist(folderResults, 'dir'), mkdir(folderResults); fprintf('Folder "%s" telah dibuat.\n', folderResults); end

try
    dataFM = readtable(fileFM);
    fprintf('Berhasil memuat data non-agregat: %s\n', fileFM);
    
    dataKW = readtable(fileKW);
    fprintf('Berhasil memuat data agregat: %s\n', fileKW);
catch ME
    errorMessage = sprintf(['Error: Tidak dapat memuat file master.\n' ...
        'Pastikan file "%s" dan "%s" ada di folder yang sama dengan skrip ini.\n' ...
        'Pesan error MATLAB: %s'], fileFM, fileKW, ME.message);
    uiwait(warndlg(errorMessage));
    return;
end

fprintf('\n--- Menjalankan Modul Interpretasi Boxplot Otomatis ---\n');
paramsToInterpret = {'FI_DT_1', 'FI_DT_3', 'FI_GV_1', 'FI_GV_3', 'FI_APS_1', 'FI_APS_2', 'FI_APS_3', 'FI_MC', 'FI_PH'};
metricResults = table('Size', [0, 7], 'VariableTypes', {'string', 'string', 'double', 'double', 'double', 'double', 'double'}, ...
                      'VariableNames', {'Grade', 'Parameter', 'IQR', 'Median', 'Jumlah_Outlier', 'Batas_Bawah', 'Batas_Atas'});
uniqueGrades = unique(dataFM.GRADE);
for i = 1:length(uniqueGrades)
    currentGrade = uniqueGrades{i};
    for j = 1:length(paramsToInterpret)
        currentParam = paramsToInterpret{j};
        
        data = dataFM.(currentParam)(strcmp(dataFM.GRADE, currentGrade));
        data(isnan(data)) = [];
        
        if ~isempty(data)
            val_iqr = iqr(data);
            val_median = median(data);
            
            quartiles = prctile(data, [25 75]);
            Q1 = quartiles(1);
            Q3 = quartiles(2);
            
            batas_atas = Q3 + 1.5 * val_iqr;
            batas_bawah = Q1 - 1.5 * val_iqr;
            
            jumlah_outlier = sum(data > batas_atas | data < batas_bawah);
            
            newRow = {string(currentGrade), string(currentParam), val_iqr, val_median, jumlah_outlier, batas_bawah, batas_atas};
            metricResults = [metricResults; newRow];
        end
    end
end

fprintf('Perhitungan metrik kunci (IQR, Median, Outlier) selesai.\n');

summaryReport = groupsummary(metricResults, 'Grade', {'mean', 'sum'}, {'IQR', 'Median', 'Jumlah_Outlier'});
summaryReport.Properties.VariableNames{'mean_Median'} = 'Rata_rata_Median';

[~, sortOrder] = sort(summaryReport.mean_IQR, 'ascend');
tempRank(sortOrder) = 1:height(summaryReport);
summaryReport.Peringkat_Konsistensi = tempRank';

[~, sortOrder] = sort(summaryReport.mean_IQR, 'descend');
tempRank(sortOrder) = 1:height(summaryReport);
summaryReport.Peringkat_Variabilitas = tempRank';

[~, sortOrder] = sort(summaryReport.sum_Jumlah_Outlier, 'descend');
tempRank(sortOrder) = 1:height(summaryReport);
summaryReport.Peringkat_Outlier = tempRank';
fprintf('Analisis agregat dan pemeringkatan grade selesai.\n');

outputFileInterpreter = fullfile(folderResults, 'Laporan_Interpretasi_Boxplot.xlsx');
writetable(summaryReport, outputFileInterpreter, 'Sheet', 'Ringkasan Peringkat Grade');
writetable(metricResults, outputFileInterpreter, 'Sheet', 'Metrik Detail per Parameter');
fprintf('Laporan interpretasi boxplot otomatis telah berhasil dibuat di: %s\n', outputFileInterpreter);

fprintf('\n--- Membuat Visualisasi Box Plot ---\n');
positions = 1:length(uniqueGrades);
fig_dt = figure('Visible', 'off', 'Position', [100 100 1000 600]);
hold on; 
h_dt(1) = plot(NaN,NaN,'-b'); 
h_dt(2) = plot(NaN,NaN,'-r');
boxplot(dataFM.FI_DT_1, dataFM.GRADE, 'Positions', positions-0.2, 'Widths', 0.3, 'Colors', 'b', 'Symbol', 'bx');
boxplot(dataFM.FI_DT_3, dataFM.GRADE, 'Positions', positions+0.2, 'Widths', 0.3, 'Colors', 'r', 'Symbol', 'rx');
hold off; 
grid on;
title('Perbandingan Distribusi FI\_DT\_1 vs FI\_DT\_3 per GRADE', 'FontSize', 14);
xlabel('GRADE', 'FontSize', 12);
ylabel('Nilai FI\_DT', 'FontSize', 12);
xticks(positions); 
xticklabels(uniqueGrades);
ylim([180 210]);
legend(h_dt, {'FI\_DT\_1', 'FI\_DT\_3'});
saveas(fig_dt, fullfile(folderPlots, 'BoxPlot_DT_Comparison.png'));
close(fig_dt);
fprintf('Plot perbandingan DT telah disimpan.\n');

fig_gv = figure('Visible', 'off', 'Position', [100 100 1000 600]);
hold on;
h_gv(1) = plot(NaN,NaN,'-b'); 
h_gv(2) = plot(NaN,NaN,'-r');
boxplot(dataFM.FI_GV_1, dataFM.GRADE, 'Positions', positions-0.2, 'Widths', 0.3, 'Colors', 'b', 'Symbol', 'bx');
boxplot(dataFM.FI_GV_3, dataFM.GRADE, 'Positions', positions+0.2, 'Widths', 0.3, 'Colors', 'r', 'Symbol', 'rx');
hold off;
grid on;
title('Perbandingan Distribusi FI\_GV\_1 vs FI\_GV\_3 per GRADE', 'FontSize', 14);
xlabel('GRADE', 'FontSize', 12);
ylabel('Nilai FI\_GV', 'FontSize', 12);
xticks(positions);
xticklabels(uniqueGrades);
ylim([180 250]);
legend(h_gv, {'FI\_GV\_1', 'FI\_GV\_3'});
saveas(fig_gv, fullfile(folderPlots, 'BoxPlot_GV_Comparison.png'));
close(fig_gv);
fprintf('Plot perbandingan GV telah disimpan.\n');

fig_aps = figure('Visible', 'off', 'Position', [100 100 1000 600]);
hold on;
h_aps(1) = plot(NaN,NaN,'-b');
h_aps(2) = plot(NaN,NaN,'-g');
h_aps(3) = plot(NaN,NaN,'-r');
boxplot(dataFM.FI_APS_1, dataFM.GRADE, 'Positions', positions-0.25, 'Widths', 0.2, 'Colors', 'b', 'Symbol', 'bx');
boxplot(dataFM.FI_APS_2, dataFM.GRADE, 'Positions', positions,      'Widths', 0.2, 'Colors', 'g', 'Symbol', 'gx');
boxplot(dataFM.FI_APS_3, dataFM.GRADE, 'Positions', positions+0.25, 'Widths', 0.2, 'Colors', 'r', 'Symbol', 'rx');
hold off;
grid on;
title('Perbandingan Distribusi FI\_APS (1, 2, 3) per GRADE', 'FontSize', 14);
xlabel('GRADE', 'FontSize', 12);
ylabel('Nilai FI\_APS', 'FontSize', 12);
xticks(positions);
xticklabels(uniqueGrades);
ylim([3 30]);
legend(h_aps, {'FI\_APS\_1', 'FI\_APS\_2', 'FI\_APS\_3'});
saveas(fig_aps, fullfile(folderPlots, 'BoxPlot_APS_Comparison.png'));
close(fig_aps);
fprintf('Plot perbandingan APS telah disimpan.\n');

params_individual = {'FI_MC', 'FI_PH'};
for i = 1:length(params_individual)
    currentParam = params_individual{i};
    fig = figure('Visible', 'off');
    boxplot(dataFM.(currentParam), dataFM.GRADE);
    title(['Distribusi Parameter ', strrep(currentParam, '_', '\_'), ' per GRADE'], 'FontSize', 14);
    xlabel('GRADE', 'FontSize', 12);
    ylabel(['Nilai ', strrep(currentParam, '_', '\_')], 'FontSize', 12);
    grid on;
    saveas(fig, fullfile(folderPlots, ['BoxPlot_', currentParam, '.png']));
    close(fig);
end
fprintf('Plot individual untuk MC dan PH telah disimpan.\n');

fprintf('\n--- Menghitung Statistik Deskriptif ---\n');
params_stats = {'FI_DT_mean', 'FI_GV_mean', 'FI_APS_mean', 'FI_MC_mean', 'FI_PH_mean'};
statDeskriptif = table();
dataKW.GRADE = upper(dataKW.GRADE);
[G, uniqueGradesAgregat] = findgroups(dataKW.GRADE);
for i = 1:length(params_stats)
    currentParam = params_stats{i};
    paramData = dataKW.(currentParam);
    
    statMean   = splitapply(@(x) mean(x, 'omitnan'), paramData, G);
    statMedian = splitapply(@(x) median(x, 'omitnan'), paramData, G);
    statStdDev = splitapply(@(x) std(x, 'omitnan'), paramData, G);
    statIQR    = splitapply(@(x) iqr(x), paramData, G);
    
    statMean   = round(statMean, 4, 'significant');
    statMedian = round(statMedian, 4, 'significant');
    statStdDev = round(statStdDev, 4, 'significant');
    statIQR    = round(statIQR, 4, 'significant');
    
    tempTable = table(uniqueGradesAgregat, repmat({currentParam}, length(uniqueGradesAgregat), 1), ...
                      statMean, statMedian, statStdDev, statIQR, ...
                      'VariableNames', {'GRADE', 'Parameter', 'Mean', 'Median', 'StdDev', 'IQR'});
    statDeskriptif = [statDeskriptif; tempTable];
end
fprintf('Perhitungan statistik deskriptif selesai.\n');

try
    outputFilePath = fullfile(folderResults, fileOutputStats);
    writetable(statDeskriptif, outputFilePath);
    fprintf('Berhasil! Tabel statistik deskriptif telah disimpan di: %s\n', outputFilePath);
catch ME
    errorMessage = sprintf(['Error: Gagal menyimpan file statistik.\n' ...
        'Pastikan Anda memiliki izin menulis di folder "%s".\n' ...
        'Pesan error MATLAB: %s'], folderResults, ME.message);
    uiwait(warndlg(errorMessage));
    return;
end
