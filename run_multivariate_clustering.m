% Conducts multivariate analysis to group grades based on their overall physicochemical profiles. 
% It first standardizes the data using a Z-score transformation. Then, it applies Hierarchical Agglomerative Clustering (HAC) with Ward's method to identify optimal clusters. 
% Finally, it uses Principal Component Analysis (PCA) to reduce dimensionality and visualize the relationships between grades and variables in a biplot, where grades are colored according to their cluster assignment from the HAC results.


inputFile = 'masterData_KW.xlsx';
outputFile_interpretasi = 'Laporan_Interpretasi_Analisis.xlsx';

if exist(outputFile_interpretasi, 'file'), delete(outputFile_interpretasi); end

try
    dataTable = readtable(inputFile);
catch ME
    errordlg(sprintf('Error: File "%s" tidak dapat ditemukan.', inputFile)); return;
end
fprintf('Berhasil memuat data.\n');

fprintf('Langkah 1 & 2: Agregasi dan Standardisasi Data...\n');
varsToAnalyze = {'FI_DT_mean', 'FI_GV_mean', 'FI_APS_mean', 'FI_MC_mean', 'FI_PH_mean'};
profileData = groupsummary(dataTable, 'GRADE', 'mean', varsToAnalyze);
meanVarNames = strcat('mean_', varsToAnalyze);
dataMatrix = profileData{:, meanVarNames};
gradeLabels = profileData.GRADE;
standardizedMatrix = zscore(dataMatrix);
fprintf('  -> Selesai.\n\n');

fprintf('Langkah 3a: Menjalankan Principal Component Analysis (PCA)...\n');
[coeff, score, latent, tsquared, explained] = pca(standardizedMatrix);
cumulativeVariance = sum(explained(1:2));
fprintf('  -> Varians kumulatif oleh PC1 & PC2: %.2f%%\n\n', cumulativeVariance);

fprintf('Langkah 3b: Menjalankan Hierarchical Clustering...\n');
euclideanDists = pdist(standardizedMatrix, 'euclidean');
linkageTree = linkage(euclideanDists, 'ward');
fprintf('  -> Pohon klaster (Linkage Ward) telah dibangun.\n\n');

fprintf('Langkah 4: Menentukan jumlah klaster optimal dan membuat visualisasi...\n');
fprintf('  -> Menentukan jumlah klaster optimal secara otomatis...\n');
N = numel(gradeLabels);
jarak_linkage = linkageTree(:, 3);
lompatan_jarak = diff(jarak_linkage);
[max_lompatan, indeks_lompatan] = max(lompatan_jarak);
k_optimal = N - indeks_lompatan;
fprintf('     - Analisis jarak Ward menemukan lompatan terbesar sebesar %.4f.\n', max_lompatan);
fprintf('     - Jumlah klaster optimal secara matematis ditentukan: k = %d\n\n', k_optimal);
clusterIDs = cluster(linkageTree, 'maxclust', k_optimal);

fig_dendro = figure('Name', 'Hasil Hierarchical Clustering', 'Position', [100, 100, 1000, 650]);
threshold = (jarak_linkage(indeks_lompatan) + jarak_linkage(indeks_lompatan + 1)) / 2;
dendrogram(linkageTree, 'Labels', gradeLabels, 'Orientation', 'top', 'ColorThreshold', threshold);
title(sprintf('Dendrogram HAC', k_optimal), 'FontSize', 16);
subtitle('Metode: Jarak Euclidean, Linkage Ward', 'FontSize', 12);
ylabel('Jarak Ward (Within-cluster variance)'); xlabel('Grade Produk');
grid on; ax = gca; ax.XTickLabelRotation = 45;

fig_elbow = figure('Name', 'Plot Siku (Elbow)', 'Position', [200, 200, 800, 600]);
plot(flipud(linkageTree(:,3)), 'o-', 'LineWidth', 1.5, 'MarkerSize', 5);
title('Plot Siku (Elbow): Jarak Penggabungan di Setiap Langkah');
xlabel('Langkah Penggabungan (dari akhir)');
ylabel('Jarak Ward (Tinggi Penggabungan)');
xlim([1, numel(gradeLabels)-1]); grid on;

variableNames = {'DT', 'GV', 'APS', 'MC', 'pH'};
fig_pca = figure('Name', 'Biplot PCA Berwarna', 'Position', [300, 300, 950, 750]);
hold on;
legendLabels = cellstr(strcat('Klaster ', num2str((1:k_optimal)')));

customColors = [
    0, 0.4470, 0.7410;
    0.8500, 0.3250, 0.0980;
    0.4660, 0.6740, 0.1880;
    0.4940, 0.1840, 0.5560;
    0.6350, 0.0780, 0.1840;
    0.3010, 0.7450, 0.9330;
];

h_scatter = gscatter(score(:,1), score(:,2), clusterIDs, customColors, 'o', 6, 'on');
for i = 1:k_optimal
    set(h_scatter(i), 'MarkerFaceColor', get(h_scatter(i), 'Color'), 'LineWidth', 1.5);
end

scale = 1.1 * max(max(abs(score(:,1:2))));
for j = 1:size(coeff, 1)
    quiver(0, 0, coeff(j,1)*scale, coeff(j,2)*scale, 0, 'r', 'LineWidth', 2, 'MaxHeadSize', 0.05);
    text(coeff(j,1)*scale*1.15, coeff(j,2)*scale*1.15, variableNames{j}, 'Color', 'r', 'FontWeight', 'bold', 'FontSize', 12);
end
hold off;

title('Biplot PCA (Diwarnai berdasarkan Klaster)', 'FontSize', 16);
xlabel(sprintf('Principal Component 1 (%.2f%%)', explained(1)));
ylabel(sprintf('Principal Component 2 (%.2f%%)', explained(2)));
grid on; axis equal;
legend(h_scatter, legendLabels, 'Location', 'northeastoutside');

fprintf('Langkah 5: Membuat laporan interpretasi di Excel...\n');
tabelScoresPCA = table(gradeLabels, score(:,1), score(:,2), clusterIDs, 'VariableNames', {'Grade', 'PC1_Score', 'PC2_Score', 'Cluster_ID'});
tabelLoadingsPCA = table(variableNames', coeff(:,1), coeff(:,2), 'VariableNames', {'Parameter', 'PC1_Loading', 'PC2_Loading'});
tabelKeanggotaan = table(gradeLabels, clusterIDs, 'VariableNames', {'Grade', 'Cluster_ID'});
tabelKeanggotaan = sortrows(tabelKeanggotaan, 'Cluster_ID');
dataDenganCluster = [profileData, table(clusterIDs, 'VariableNames', {'Cluster_ID'})];
profilMean = groupsummary(dataDenganCluster, 'Cluster_ID', 'mean', meanVarNames);
tempCount = groupsummary(dataDenganCluster, 'Cluster_ID');
profilCount = tempCount(:, {'Cluster_ID', 'GroupCount'});
profilCount.Properties.VariableNames{'GroupCount'} = 'Jumlah_Anggota';
profilCluster = join(profilMean, profilCount);
contohGrades = cell(k_optimal, 1);
for i = 1:k_optimal
    gradesInCluster = tabelKeanggotaan.Grade(tabelKeanggotaan.Cluster_ID == i);
    contohGrades{i} = strjoin(gradesInCluster(1:min(3, end))', ', ');
end
profilCluster.Contoh_Grade = contohGrades;
profilCluster = movevars(profilCluster, 'Contoh_Grade', 'After', 'Jumlah_Anggota');
profilCluster = movevars(profilCluster, 'Jumlah_Anggota', 'After', 'Cluster_ID');
writetable(tabelScoresPCA, outputFile_interpretasi, 'Sheet', 'PCA_Scores_per_Cluster');
writetable(tabelLoadingsPCA, outputFile_interpretasi, 'Sheet', 'PCA_Loadings');
writetable(tabelKeanggotaan, outputFile_interpretasi, 'Sheet', 'Cluster_Keanggotaan');
writetable(profilCluster, outputFile_interpretasi, 'Sheet', 'Cluster_Profil');

fprintf('\nLangkah 6: Menyimpan semua hasil visualisasi...\n');
try
    print(fig_pca, 'Biplot_PCA_berwarna_otomatis.png', '-dpng', '-r300');
    print(fig_dendro, 'Dendrogram_HAC_Ward_otomatis.png', '-dpng', '-r300');
    print(fig_elbow, 'Plot_Siku_Penentuan_Cluster.png', '-dpng', '-r300');
    fprintf('  -> Semua gambar visualisasi (Biplot, Dendrogram, Plot Siku) telah disimpan.\n');
    close(fig_pca, fig_dendro, fig_elbow);
catch ME
    warndlg('Gagal menyimpan gambar.');
end
