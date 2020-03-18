[file, dir] = uigetfile({'*.xlsx'; '*.xls'; '*.ods'});
if ~file
    return
end
path = fullfile(dir,file);
kopts = detectImportOptions(path, 'Sheet', 'Keywords', 'ReadVariableNames', false);
keywords = string(readmatrix(path, kopts));

Tw = readtable(path, detectImportOptions(path, 'Sheet', 'Workers'));
Tp = readtable(path, detectImportOptions(path, 'Sheet', 'Projects'));

P = zeros(height(Tp), length(keywords));

quota_sum = 0;
for i = 1:height(Tw)
    quota_sum = quota_sum + Tw{i,2};
end
W = zeros(quota_sum, length(keywords));

if quota_sum < height(Tp)
    warning("Sum of quotas is less than number of projects. Not all projects will be allocated.");
end

for i = 1:height(Tp)
    for j = 3:2:(width(Tp)-2)
        keyword = string(Tp{i,j});
        if keyword == ""
            continue
        end
        P(i, find(keywords == keyword)) = Tp{i, j+1};
    end
end

worker_names = cell(quota_sum, 1);
n = 1;
for i = 1:height(Tw)
    q = Tw{i,2};
    for j = 3:2:(width(Tw)-1)
        keyword = string(Tw{i,j});
        if keyword == ""
            continue
        end
        W(n:n+q-1, find(keywords == keyword)) = Tw{i, j+1};
    end
    worker_names(n:n+q-1) = Tw{i,1};
    n = n + q;
end

W = W ./ vecnorm(W,2,2);
P = P ./ vecnorm(P,2,2);

S = W * P';

[assignment, cost] = assignmentoptimal(1 - S);

project_names = cell(length(assignment), 1);
for i = 1:length(assignment)
    project_names{i} = Tp{assignment(i), 1};
end
project_names = string(project_names);
worker_names = string(worker_names);
scores = S(sub2ind(size(S), 1:quota_sum, assignment'))';

T = table(worker_names, project_names, scores);
T.Properties.VariableNames = {'Worker'; 'Project'; 'Score'};

[file, dir] = uiputfile({'*.xlsx'; '*.csv'});
if ~file
    return
end
writetable(T, fullfile(dir, file));