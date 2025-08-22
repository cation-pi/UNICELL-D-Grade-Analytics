% It evaluates product performance against specifications, and generates a ranked report of the grades based on their consistency metrics.

fileData = 'StatDeskriptif_Agregat.xlsx';
fileSpec = 'Spesifikasi.xlsx';
fileOutput = 'Hasil_Evaluasi_Kinerja.xlsx';

fprintf('Memulai proses evaluasi kinerja produk (Versi Final - Perbaikan)...\n');

try
    dataTbl = readtable(fileData);
    fprintf('Berhasil membaca file data: %s\n', fileData);
    
    specTblRaw = readtable(fileSpec);
    fprintf('Berhasil membaca file spesifikasi: %s\n', fileSpec);
catch ME
    errorMessage = sprintf('Error saat membaca file: %s. Pastikan file berada di folder yang sama.', ME.message);
    error(errorMessage);
end

specTblRaw.Properties.VariableNames = strtrim(specTblRaw.Properties.VariableNames);

paramMap = {
    'FI_DT_mean',  'DT';
    'FI_GV_mean',  'GV';
    'FI_MC_mean',  'MC';
    'FI_PH_mean',  'PH';
    'FI_APS_mean', 'APS'
};

tidySpecData = cell(0, 4);
fprintf('Memproses dan menata ulang tabel spesifikasi...\n');

for i = 1:height(specTblRaw)
    gradeName = specTblRaw.GradeADCA{i};
    if isempty(gradeName), continue; end
    
    for j = 1:size(paramMap, 1)
        paramNameInData = paramMap{j, 1};
        colNameInSpec = paramMap{j, 2};
        
        if ~ismember(colNameInSpec, specTblRaw.Properties.VariableNames), continue; end
        
        specValue = specTblRaw.(colNameInSpec)(i);
        
        if iscell(specValue)
            specStr = specValue{1};
        elseif isstring(specValue)
            specStr = char(specValue);
        elseif isnumeric(specValue)
            specStr = num2str(specValue);
        else
            specStr = '';
        end
        
        [minVal, maxVal] = parseSpecString(specStr);
        tidySpecData = [tidySpecData; {gradeName, paramNameInData, minVal, maxVal}];
    end
end

specTbl = cell2table(tidySpecData, 'VariableNames', {'GRADE', 'Parameter', 'min', 'max'});
fprintf('Tabel spesifikasi berhasil diproses.\n');

uniqueParams = unique(dataTbl.Parameter);
numParams = length(uniqueParams);
if exist(fileOutput, 'file'), delete(fileOutput); end

fprintf('Memulai evaluasi untuk %d parameter...\n', numParams);
for i = 1:numParams
    currentParam = uniqueParams{i};
    if ~ismember(currentParam, paramMap(:,1)), continue; end
    
    fprintf('Memproses parameter: %s (%d/%d)...\n', currentParam, i, numParams);
    
    paramData = dataTbl(strcmp(dataTbl.Parameter, currentParam), :);
    paramSpec = specTbl(strcmp(specTbl.Parameter, currentParam), :);
    
    mergedTbl = outerjoin(paramData, paramSpec, 'Keys', 'GRADE', 'Type', 'left', 'MergeKeys', true);
    
    numRows = height(mergedTbl);
    rangeCol = cell(numRows, 1);
    statusCol = strings(numRows, 1);
    
    for j = 1:numRows
        minSpec = mergedTbl.min(j);
        maxSpec = mergedTbl.max(j);
        
        if isnan(minSpec) || isnan(maxSpec)
            rangeCol{j} = 'Tidak Ditemukan';
            statusCol(j) = "Spesifikasi Tidak Ditemukan";
            continue;
        end
        
        rangeCol{j} = sprintf('%.2f–%.2f', minSpec, maxSpec);
        
        meanVal = mergedTbl.Mean(j);
        medianVal = mergedTbl.Median(j);
        
        if (meanVal >= minSpec && meanVal <= maxSpec) && (medianVal >= minSpec && medianVal <= maxSpec)
            statusCol(j) = "Memenuhi";
        else
            statusCol(j) = "Tidak Memenuhi";
        end
    end
    
    mergedTbl.('Rentang Spesifikasi') = rangeCol;
    mergedTbl.('Status Kinerja') = statusCol;
    
    resultTbl = mergedTbl(:, {'GRADE', 'Mean', 'Median', 'StdDev', 'IQR', 'Rentang Spesifikasi', 'Status Kinerja'});
    
    writetable(resultTbl, fileOutput, 'Sheet', currentParam);
end

fprintf('\nProses evaluasi selesai.\n');
fprintf('Hasil telah berhasil disimpan di file: %s\n', fileOutput);

function [minVal, maxVal] = parseSpecString(specStr)
    minVal = NaN; maxVal = NaN;
    if isempty(specStr) || (isstring(specStr) && strlength(specStr) == 0), return; end
    
    s = strrep(specStr, ',', '.');
    s = strtrim(s);
    
    if contains(s, '-')
        parts = strsplit(s, '-');
        if numel(parts) == 2, minVal = str2double(parts{1}); maxVal = str2double(parts{2}); end
    elseif startsWith(s, '<')
        val = str2double(regexprep(s, '<', ''));
        minVal = 0; maxVal = val;
    elseif startsWith(s, '>')
        val = str2double(regexprep(s, '>', ''));
        minVal = val; maxVal = inf;
    else
        val = str2double(s);
        if ~isnan(val), minVal = val; maxVal = val; end
    end
end
