% % Prob 4
% h = 400*1000;
% r = 6378.1*1000;
% tecu = 80;
% el = 30; % deg
% c = 3*10^8; %m/s
% s = 1/(cosd(asind(r/(r+h)*cosd(el))));
% 
% vertical_tec = tecu/s % Vertical TECu


new_f_muf = 15*10^6;
newLaunchAngle = 20; % degrees
iono = (2*10^6)*10^6; % m^-3
q = 1.6*10^-19;
m = 9.1*10^-31;
epsilon = 8.85*10^-12;
new_f_0 = new_f_muf/(secd(90-newLaunchAngle));
new_omega = 2*pi*new_f_0;
nmf2 = ((new_omega)^2)*m*epsilon/q/q

% iono = (2*10^6)*10^6; % m^-3
% q = 1.6*10^-19;
% m = 9.1*10^-31;
% epsilon = 8.85*10^-12;
% omega_2 = iono*(q^2)/m/epsilon;
% f_0 = sqrt(omega_2)/2/pi;
% f_muf = (f_0/10^6)*secd(90-launchAngle) % MHz

c = 3*10^8;
rgc = 1200;
virtual_height = (rgc/2)*tand(newLaunchAngle) 
rs = 2*(rgc/2)/cosd(newLaunchAngle) 
t = rs*1000/c/(10^-3)
