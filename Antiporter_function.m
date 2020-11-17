function [dAAindt]=Antiporter_function(AAex,AAin,Kmex,Kmin,rVmax,Vmax,dt,Vol)

% Apparent Km
KappAAex=zeros(1,22);
KappAAin=zeros(1,22);
Kmex(isnan(Kmex))=inf;
Kmin(isnan(Kmin))=inf;
rVmax(isnan(rVmax))=0;
Xex=zeros(1,22);
Xin=zeros(1,22);

for i=1:22
    
    for n=1:22
        Xex(n)=(AAex(n)/Kmex(n));
        Xin(n)=(AAin(n)/Kmin(n));
    end
    
    KappAAex(i)=Kmex(i)*(1+sum(Xex)-(AAex(i)/Kmex(i)));
    KappAAin(i)=Kmin(i)*(1+sum(Xin)-(AAin(i)/Kmin(i)));
end

% Fractional Saturation
FSex=zeros(1,22);
FSin=zeros(1,22);

for i=1:22
    FSex(i)=AAex(i)/(KappAAex(i)+AAex(i));
    FSin(i)=AAin(i)/(KappAAin(i)+AAin(i));
end

% Total fractional saturation
Ftex=sum(FSex);
Ftin=sum(FSin);

% Total flux
Jtoi=(Vmax*Ftex*Ftin)*dt;
Jtio=Jtoi;

% Individual  Flux
Joi=zeros(1,22);
Jio=zeros(1,22);

for i=1:22
    if Ftex==0
        Ftex=inf;
    end
    if Ftin==0
        Ftin=inf;
    end
    Joi(i)=Jtoi*rVmax(i)*(FSex(i)/Ftex);
    Jio(i)=Jtio*rVmax(i)*(FSin(i)/Ftin);
    if (isnan(Joi(i)))&&(isnan(Jio(i)))
        Joi(i)=0;
        Jio(i)=0;
    end
end

% Change in intracellular concentration for each AA
dAAindt=zeros(1,22);

for i=1:22
    dAAindt(i)= (Joi(i)-Jio(i))/Vol;
end

end