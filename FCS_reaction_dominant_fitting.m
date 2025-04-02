clc;
clear;

filename = '1.xlsx';
col_names = readcell(filename, 'Range', '1:1');
data = readmatrix(filename, 'Range', 'A2');

x_col = 3;  
y_cols = 4:2:size(data,2);  


for y_col = y_cols
    try
        current_col_name = col_names{y_col};
        if isnumeric(current_col_name)
            current_col_name = num2str(current_col_name);
        else
            current_col_name = char(current_col_name);
        end
        valid_filename = regexprep(current_col_name, '[^a-zA-Z0-9_-]', '_');
        
       
        x_data = data(:, x_col);
        y_data = data(:, y_col);
        
        x_clean = rmmissing(x_data);
        y_clean = rmmissing(y_data);
        
        minlen = min(length(x_clean), length(y_clean));
        if minlen < 2
            fprintf('Skipping column %d: Insufficient data\n', y_col);
            continue;
        end
        
        x = x_clean(1:minlen)';
        y = y_clean(1:minlen)';
        
        
        B = 0.7;
        koff0 = 1;
        initialGuess = [40, 0.1, koff0*B/(1-B), koff0];
        lb = [0, 0, 0, 0];
        ub = [inf, inf, inf, inf];
        options = optimoptions('lsqcurvefit', 'Display', 'off', 'TolFun', 1e-9);
        
        % Curve fitting
        [params, resnorm] = lsqcurvefit(@modelFcn, initialGuess, x, y, lb, ub, options);
         B_value = params(3) / (params(3) + params(4));  % kon/(kon+koff)
        
        % Generate fitted curve
        y_fit = modelFcn(params, x);
        
        % Plot the graph
        figure;
        scatter(x, y, 80, 'Marker', 'o', 'MarkerEdgeColor', [0.5 0.5 0.5], 'MarkerFaceColor', 'none', 'LineWidth', 0.5);    
        hold on;
        plot(x, y_fit, 'r-', 'LineWidth', 2);
        title(sprintf('Fitting Result - %s', current_col_name));
        xlabel('X');
        ylabel('Y');
        legend('Original Data', 'Fitted Curve', 'Location', 'best');
        
        % Save as EPS file
        saveas(gcf, [valid_filename '.eps'], 'epsc');
        close(gcf);
        
        % Display fitting information
        fprintf('Column Name: %s\n', current_col_name);
        fprintf('Fitting Parameters: [%.4f, %.4f, %.4f, %.4f]\n', params);
        fprintf('B value: %.6f\n', B_value);  
        fprintf('Residual Sum of Squares: %.4f\n\n', resnorm);
        
    catch ME
        fprintf('Error processing column %d: %s\n', y_col, ME.message);
        continue;
    end
end

function y_fit = modelFcn(params, x)
    E = log(10);
    N = params(1);
    D = params(2);
    kon = params(3);
    koff = params(4);
    
    F = koff / (kon + koff);
    C = kon / (kon + koff);
    X1 = exp(E * x) / 1000;
    
    term1 = F ./ (sqrt(1 + (4*X1*D)/(0.291^2)) .* ...
                  sqrt(1 + (4*X1*D)/(0.291^2)) .* ...
                  sqrt(1 + (4*X1*D)/(2.225^2)));
    term2 = C .* exp(-koff * X1);
    
    y_fit = (1/(N*(2^1.5))) * (term1 + term2);
end
