% This script specifically evaluates the internal consistency within each product grade. 
% It statistically compares measurements from different sampling points (top, middle, bottom). 
% It uses the Wilcoxon Rank-Sum (Mann-Whitney U) test for parameters with two sampling points and the Kruskal-Wallis test for parameters with three, ultimately classifying each grade as "Homogen" or "Not Homogen".

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

fprintf('\n--- Menjalankan Modul Analisis Homogenitas Internal ---\n');

paramGroups = {
    {'DT',  {'FI_DT_1', 'FI_DT_3'}}, ...
    {'GV',  {'FI_GV_1', 'FI_GV_3'}}, ...
    {'APS', {'FI_APS_1', 'FI_APS_2', 'FI_APS_3'}}
};

homogeneityResults = table('Size', [0, 5], 'VariableTypes', {'string', 'string', 'double', 'double', 'string'}, ...
                           'VariableNames', {'Grade', 'Parameter_Utama', 'Delta_Median', 'p_value', 'Kesimpulan_Homogenitas'});

uniqueGrades = unique(dataFM.GRADE);

for i = 1:length(uniqueGrades)
    currentGrade = uniqueGrades{i};
    
    for j = 1:length(paramGroups)
        paramName = paramGroups{j}{1};
        paramCols = paramGroups{j}{2};
        
        gradeData = dataFM(strcmp(dataFM.GRADE, currentGrade), :);
        
        if length(paramCols) == 2
            data1 = gradeData.(paramCols{1}); data1(isnan(data1)) = [];
            data2 = gradeData.(paramCols{2}); data2(isnan(data2)) = [];
            
            if ~isempty(data1) && ~isempty(data2)
                deltaMedian = abs(median(data1) - median(data2));
                p_value = ranksum(data1, data2);
            else
                deltaMedian = NaN; p_value = NaN;
            end
            
        elseif length(paramCols) == 3
            data1 = gradeData.(paramCols{1}); data1(isnan(data1)) = [];
            data2 = gradeData.(paramCols{2}); data2(isnan(data2)) = [];
            data3 = gradeData.(paramCols{3}); data3(isnan(data3)) = [];
            
            if ~isempty(data1) && ~isempty(data2) && ~isempty(data3)
                medians = [median(data1), median(data2), median(data3)];
                deltaMedian = max(medians) - min(medians);
                
                data_vector = [data1; data2; data3];
                group_vector = [ones(size(data1)); 2*ones(size(data2)); 3*ones(size(data3))];
                p_value = kruskalwallis(data_vector, group_vector, 'off');
            else
                deltaMedian = NaN; p_value = NaN;
            end
        end
        
        if isnan(p_value)
            kesimpulan = 'Data Tidak Cukup';
        elseif p_value < 0.05
            kesimpulan = 'Tidak Homogen Signifikan';
        else
            kesimpulan = 'Homogen';
        end
        
        newRow = {string(currentGrade), string(paramName), deltaMedian, p_value, string(kesimpulan)};
        homogeneityResults = [homogeneityResults; newRow];
    end
end

fprintf('Analisis statistik homogenitas (Uji Wilcoxon & Kruskal-Wallis) selesai.\n');

[G, grades] = findgroups(homogeneityResults.Grade);

skorInhomogenitas = splitapply(@(x) sum(strcmp(x, 'Tidak Homogen Signifikan')), homogeneityResults.Kesimpulan_Homogenitas, G);

avgDeltaMedian = splitapply(@(x) mean(x, 'omitnan'), homogeneityResults.Delta_Median, G);
summaryHomogeneity = table(grades, skorInhomogenitas, avgDeltaMedian, 'VariableNames', {'Grade', 'Skor_Inhomogenitas', 'Rata_Rata_Delta_Median'});

summaryHomogeneity = sortrows(summaryHomogeneity, 'Skor_Inhomogenitas', 'descend');
fprintf('Pemeringkatan grade berdasarkan skor inhomogenitas selesai.\n');

outputFileHomogeneity = fullfile(folderResults, 'Laporan_Analisis_Homogenitas.xlsx');
writetable(summaryHomogeneity, outputFileHomogeneity, 'Sheet', 'Peringkat Inhomogenitas');
writetable(homogeneityResults, outputFileHomogeneity, 'Sheet', 'Detail Analisis Homogenitas');
fprintf('Laporan analisis homogenitas internal otomatis telah berhasil dibuat di: %s\n', outputFileHomogeneity);
