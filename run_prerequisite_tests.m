% This script is dedicated to performing prerequisite statistical tests on the aggregate dataset (masterData_KW.xlsx). 
% It automatically selects and runs normality tests (Shapiro-Wilk for n<50, Kolmogorov-Smirnov for n≥50) and homogeneity of variance tests (Levene's or Brown-Forsythe). 
% The results from this script are crucial for validating the choice of non-parametric methods for the main analysis.


fileKW = 'masterData_KW.xlsx';
params_to_test = {'FI_DT_mean', 'FI_GV_mean', 'FI_APS_mean', 'FI_MC_mean', 'FI_PH_mean'};

try
    dataKW = readtable(fileKW);
    fprintf('Berhasil memuat data agregat: %s\n', fileKW);
catch ME
    errorMessage = sprintf(['Error: Tidak dapat memuat file master.\n' ...
        'Pastikan file "%s" ada di folder yang sama dengan skrip ini.\n' ...
        'Pesan error MATLAB: %s'], fileKW, ME.message);
    uiwait(warndlg(errorMessage));
    return;
end

if ~exist('swtest.m', 'file')
    errorMsg = ['Fungsi "swtest.m" tidak ditemukan. \n\n' ...
        'Harap unduh dari MATLAB File Exchange (ID: 13964) '...
        'dan letakkan di folder yang sama dengan skrip ini.'];
    errordlg(errorMsg, 'Dependensi Eksternal Hilang');
    return;
end

fprintf('Info: Fungsi swtest.m untuk uji Shapiro-Wilk ditemukan.\n\n');

ujiNormalitas = table();
indikatorDistribusi = table();
ujiHomogenitas = table();

for i = 1:length(params_to_test)
    currentParam = params_to_test{i};
    fprintf('--- Menganalisis Parameter: %s ---\n', currentParam);
    any_non_normal = false;
    any_outliers = false;
    any_skewed = false;
    sample_sizes = [];
    
    paramData = dataKW.(currentParam);
    paramGroup = dataKW.GRADE;
    validDataIdx = ~isnan(paramData);
    paramData = paramData(validDataIdx);
    paramGroup = paramGroup(validDataIdx);
    
    uniqueGrades = unique(paramGroup);
    
    for j = 1:length(uniqueGrades)
        currentGrade = uniqueGrades{j};
        
        gradeData = paramData(strcmp(paramGroup, currentGrade));
        n = length(gradeData);
        sample_sizes = [sample_sizes, n];
        
        if n < 3
            continue;
        end
        
        if n < 50
            metodeUji = 'Shapiro-Wilk';
            [~, pValue] = swtest(gradeData);
        else
            metodeUji = 'Kolmogorov-Smirnov';
            [~, pValue] = lillietest(gradeData);
        end
        
        isNormal = pValue >= 0.05;
        if ~isNormal
            any_non_normal = true;
        end
        
        tempNorm = table({currentGrade}, {currentParam}, n, {metodeUji}, round(pValue, 4, 'significant'), ...
            iif(isNormal, "Ya", "Tidak"), 'VariableNames', {'GRADE', 'Parameter', 'n', 'MetodeUji', 'pValue', 'Normal'});
        ujiNormalitas = [ujiNormalitas; tempNorm];
        
        hasOutlier = any(isoutlier(gradeData));
        if hasOutlier
            any_outliers = true;
        end
        
        skewVal = skewness(gradeData);
        if abs(skewVal) > 1
            any_skewed = true;
        end
        
        tempDist = table({currentGrade}, {currentParam}, iif(hasOutlier, "Ya", "Tidak"), round(skewVal, 4, 'significant'), ...
            'VariableNames', {'GRADE', 'Parameter', 'AdaOutlier', 'Skewness'});
        indikatorDistribusi = [indikatorDistribusi; tempDist];
    end
    
    is_unbalanced = (max(sample_sizes) / min(sample_sizes)) > 1.5;
    if any_non_normal || any_outliers || any_skewed || is_unbalanced
        metodeHomogen = 'Brown-Forsythe';
        testType = 'BrownForsythe';
    else
        metodeHomogen = 'Levene';
        testType = 'LeveneAbsolute';
    end
    
    pHomogen = vartestn(paramData, paramGroup, 'TestType', testType, 'Display', 'off');
    
    keputusan = iif(pHomogen < 0.05, "Tolak H0 (Varians tidak homogen)", "Gagal Tolak H0 (Varians homogen)");
    
    tempHomogen = table({currentParam}, {metodeHomogen}, round(pHomogen, 4, 'significant'), {keputusan}, ...
        'VariableNames', {'Parameter', 'MetodeUji', 'pValue', 'Keputusan'});
    ujiHomogenitas = [ujiHomogenitas; tempHomogen];
    fprintf('Uji Homogenitas (%s) selesai dengan p-value = %.4f\n\n', metodeHomogen, pHomogen);
end

try
    writetable(ujiNormalitas, 'Hasil_Uji_Normalitas.xlsx');
    fprintf('Berhasil menyimpan: Hasil_Uji_Normalitas.xlsx\n');
    
    writetable(indikatorDistribusi, 'Hasil_Indikator_Distribusi.xlsx');
    fprintf('Berhasil menyimpan: Hasil_Indikator_Distribusi.xlsx\n');
    
    writetable(ujiHomogenitas, 'Hasil_Uji_Homogenitas.xlsx');
    fprintf('Berhasil menyimpan: Hasil_Uji_Homogenitas.xlsx\n');
catch ME
    errorMessage = sprintf(['Error: Gagal menyimpan file hasil.\n' ...
        'Pastikan tidak ada file hasil yang sedang terbuka di Excel.\n' ...
        'Pesan error MATLAB: %s'], ME.message);
    uiwait(warndlg(errorMessage));
    return;
end

fprintf('\nSkrip uji prasyarat telah selesai dijalankan.\n');

function C = iif(condition, A, B)
    if condition
        C = A;
    else
        C = B;
    end
end
