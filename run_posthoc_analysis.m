% This script is executed following a significant result from the Kruskal-Wallis test. 
% It performs a Dunn's post-hoc test with a Bonferroni correction to identify which specific pairs of ADCA grades are significantly different from each other. 
% The output includes detailed comparison tables, p-value heatmaps, and a network graph to visualize the significant relationships.


inputFile = 'masterData_KW.xlsx';
outputFile = 'Laporan_PostHoc_Dunn_Bonferroni.xlsx';
outputDetailFile = 'Laporan_Detail_Signifikansi.xlsx';
outputImageFolder = 'Heatmap_Results';
outputSummaryFolder = 'Summary_Visuals';

if exist(outputFile, 'file'), delete(outputFile); end
if exist(outputDetailFile, 'file'), delete(outputDetailFile); end
if exist(outputImageFolder, 'dir'), rmdir(outputImageFolder, 's'); end
if exist(outputSummaryFolder, 'dir'), rmdir(outputSummaryFolder, 's'); end
mkdir(outputImageFolder);
mkdir(outputSummaryFolder);
fprintf('Info: File laporan dan folder gambar lama telah dibersihkan.\n');

try
    dataTable = readtable(inputFile);
    fprintf('Berhasil memuat data dari: %s\n\n', inputFile);
catch ME
    errorMessage = sprintf('Error: File "%s" tidak dapat ditemukan.\n', inputFile);
    errordlg(errorMessage, 'File Tidak Ditemukan');
    return;
end

dependentVars = {'FI_DT_mean', 'FI_GV_mean', 'FI_APS_mean', 'FI_MC_mean', 'FI_PH_mean'};
paramAbbr = {'DT', 'GV', 'APS', 'MC', 'PH'};
alpha = 0.05;
dataTable = dataTable(~isundefined(categorical(dataTable.GRADE)), :);
groupNames = unique(dataTable.GRADE, 'stable');
numGroups = numel(groupNames);
numParams = numel(dependentVars);
significanceCube = zeros(numGroups, numGroups, numParams);
fprintf('Analisis akan dilakukan untuk %d parameter.\n\n', numParams);

for i = 1:numParams
    currentVar = dependentVars{i};
    fprintf('Memproses: %s (%d dari %d)...\n', currentVar, i, numParams);
    
    data = dataTable.(currentVar);
    group = dataTable.GRADE;
    validIdx = ~isnan(data);
    data = data(validIdx);
    group = group(validIdx);
    
    [~, ~, stats] = kruskalwallis(data, group, 'off');
    comparisonMatrix = multcompare(stats, 'Estimate', 'kruskalwallis', 'CType', 'bonferroni', 'Display', 'off');
    
    resultsTable = array2table(comparisonMatrix, 'VariableNames', ...
        {'Grup_1_Idx', 'Grup_2_Idx', 'Lower_CI', 'Estimate', 'Upper_CI', 'Adjusted_p_value'});
    
    for k = 1:height(resultsTable)
        if resultsTable.Adjusted_p_value(k) < alpha
            idx1 = resultsTable.Grup_1_Idx(k);
            idx2 = resultsTable.Grup_2_Idx(k);
            significanceCube(idx1, idx2, i) = 1;
            significanceCube(idx2, idx1, i) = 1;
        end
    end
    
    resultsTable.Grup_1 = groupNames(comparisonMatrix(:,1));
    resultsTable.Grup_2 = groupNames(comparisonMatrix(:,2));
    formattedTable = resultsTable(:, {'Grup_1', 'Grup_2', 'Lower_CI', 'Estimate', 'Upper_CI', 'Adjusted_p_value'});
    formattedTable.Signifikan = strings(height(formattedTable), 1);
    significantRowsIdx = formattedTable.Adjusted_p_value < alpha;
    formattedTable.Signifikan(significantRowsIdx) = "*";
    sheetName = ['Dunn_' currentVar(4:end-5)];
    writetable(formattedTable, outputFile, 'Sheet', sheetName);
    
    pValueMatrix = ones(numGroups);
    for k = 1:height(resultsTable)
        idx1 = resultsTable.Grup_1_Idx(k);
        idx2 = resultsTable.Grup_2_Idx(k);
        pValueMatrix(idx1, idx2) = resultsTable.Adjusted_p_value(k);
        pValueMatrix(idx2, idx1) = resultsTable.Adjusted_p_value(k);
    end
    
    cellLabels = strings(numGroups);
    for r = 1:numGroups
        for c = 1:numGroups
            if r ~= c && pValueMatrix(r, c) < alpha
                cellLabels(r, c) = "*";
            end
        end
    end
    
    fig = figure('Visible', 'off', 'Position', [100 100 850 700]);
    imagesc(pValueMatrix);
    colormap(parula);
    colorbarHandle = colorbar;
    ylabel(colorbarHandle, 'Adjusted P-Value');
    clim([0 1]);
    ax = gca;
    ax.XTick = 1:numGroups;
    ax.YTick = 1:numGroups;
    ax.XTickLabel = groupNames;
    ax.YTickLabel = groupNames;
    ax.XTickLabelRotation = 45;
    
    title(sprintf('Heatmap P-Value (Dunn-Bonferroni) untuk: %s', strrep(currentVar, '_', '\_')), 'FontSize', 14);
    
    for r = 1:numGroups
        for c = 1:numGroups
            textColor = iif(pValueMatrix(r,c) > 0.6, 'black', 'white');
            text(c, r, cellLabels(r, c), 'HorizontalAlignment', 'center', 'Color', textColor, 'FontSize', 14, 'FontWeight', 'bold');
        end
    end
    
    outputImageName = fullfile(outputImageFolder, ['Heatmap_Dunn_' currentVar(4:end-5) '.png']);
    print(fig, outputImageName, '-dpng', '-r300');
    close(fig);
    
    fprintf('  -> Laporan Excel dan Heatmap untuk %s telah dibuat.\n\n', currentVar);
end

fprintf('\n------------------------------------------------------------\n');
fprintf('Semua analisis per parameter selesai.\nMemulai pembuatan laporan dan visualisasi agregat...\n');

frequencyMatrix = sum(significanceCube, 3);

fprintf('  -> Membuat Heatmap Frekuensi Signifikansi...\n');
fig_freq = figure('Visible', 'off', 'Position', [100 100 900 750]);
imagesc(frequencyMatrix);
colormap(parula);
colorbarHandle = colorbar;
ylabel(colorbarHandle, sprintf('Jumlah Parameter Signifikan (dari %d)', numParams));
clim([0 numParams]);
ax = gca;
ax.XTick = 1:numGroups;
ax.YTick = 1:numGroups;
ax.XTickLabel = groupNames;
ax.YTickLabel = groupNames;
ax.XTickLabelRotation = 45;
title('Heatmap Frekuensi Signifikansi Antar Grade', 'FontSize', 14);
for r = 1:numGroups
    for c = 1:numGroups
        score = frequencyMatrix(r, c);
        textColor = iif(score > numParams * 0.6, 'black', 'white');
        if r ~=c
             text(c, r, num2str(score), 'HorizontalAlignment', 'center', 'Color', textColor, 'FontSize', 10, 'FontWeight', 'bold');
        end
    end
end
outputImageName_Freq = fullfile(outputSummaryFolder, 'Heatmap_Agregat_Frekuensi.png');
print(fig_freq, outputImageName_Freq, '-dpng', '-r300');
close(fig_freq);

fprintf('  -> Membuat Graf Jaringan Hubungan Paling Konsisten...\n');
adjacencyMatrix = (frequencyMatrix == numParams);
if ~any(adjacencyMatrix, 'all')
    fprintf('  -> Info: Tidak ada pasangan grade yang signifikan di seluruh %d parameter. Graf Jaringan tidak dibuat.\n', numParams);
else
    G = graph(adjacencyMatrix, groupNames, 'upper');
    fig_graph = figure('Visible', 'off', 'Position', [100 100 800 800]);
    p = plot(G, 'Layout', 'circle', 'NodeColor', '#0072BD', 'EdgeColor', '#D95319', 'LineWidth', 3, 'MarkerSize', 12, 'NodeLabel', {});
    text(p.XData, p.YData, G.Nodes.Name, 'FontSize', 10, 'HorizontalAlignment','center', 'VerticalAlignment','bottom');
    title(sprintf('Graf Jaringan: Pasangan Grade Signifikan di Seluruh %d Parameter', numParams), 'FontSize', 14);
    set(gca, 'XTick', [], 'YTick', []);
    box on;
    outputImageName_Graph = fullfile(outputSummaryFolder, 'Graf_Jaringan_Signifikansi_Total.png');
    print(fig_graph, outputImageName_Graph, '-dpng', '-r300');
    close(fig_graph);
end

fprintf('  -> Membuat Laporan Excel Detail Signifikansi...\n');
detailSignifikanTabel = table('Size', [0, 4], ...
    'VariableTypes', {'string', 'string', 'double', 'string'}, ...
    'VariableNames', {'Grup_1', 'Grup_2', 'Jumlah_Signifikan', 'Parameter_Signifikan'});
for r = 1:numGroups
    for c = r+1:numGroups
        jumlahSignifikan = frequencyMatrix(r, c);
        if jumlahSignifikan > 0
            grup1 = string(groupNames{r});
            grup2 = string(groupNames{c});
            signifVector = squeeze(significanceCube(r, c, :));
            signifIndices = find(signifVector);
            paramList = paramAbbr(signifIndices);
            paramString = strjoin(paramList, ', ');
            newRow = {grup1, grup2, jumlahSignifikan, paramString};
            detailSignifikanTabel = [detailSignifikanTabel; newRow];
        end
    end
end
detailSignifikanTabel = sortrows(detailSignifikanTabel, 'Jumlah_Signifikan', 'descend');
writetable(detailSignifikanTabel, outputDetailFile);

fprintf('\n------------------------------------------------------------\n');
fprintf('Pipeline analisis final telah selesai.\n\n');
fprintf('Output yang dihasilkan:\n');
fprintf('1. Laporan Excel Post-Hoc: %s\n', outputFile);
fprintf('2. Laporan Excel Detail Signifikansi: %s\n', outputDetailFile);
fprintf('3. Heatmap per parameter di folder: %s\n', outputImageFolder);
fprintf('4. Visualisasi kesimpulan di folder: %s\n', outputSummaryFolder);
fprintf('------------------------------------------------------------\n');

function C = iif(condition, A, B)
    if condition
        C = A;
    else
        C = B;
    end
end
