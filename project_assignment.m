% Copyright (c) 2020, Alexander Huntley
% All rights reserved.
% 
% Redistribution and use in source and binary forms, with or without
% modification, are permitted provided that the following conditions are
% met:
% 
%     * Redistributions of source code must retain the above copyright
%       notice, this list of conditions and the following disclaimer.
%     * Redistributions in binary form must reproduce the above copyright
%       notice, this list of conditions and the following disclaimer in
%       the documentation and/or other materials provided with the distribution
% 
% THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
% AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
% IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
% ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
% LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
% CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
% SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
% INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
% CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
% ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
% POSSIBILITY OF SUCH DAMAGE.

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

[file, dir] = uiputfile({'*.xlsx'; '*.xls'; '*.dat'; '*.txt'; '*.csv'});
if ~file
    return
end
writetable(T, fullfile(dir, file));