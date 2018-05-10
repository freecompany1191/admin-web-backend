package com.o2osys.mng;

import java.util.Locale;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.PropertySource;
import org.springframework.context.annotation.PropertySources;
import org.springframework.context.support.ReloadableResourceBundleMessageSource;
import org.springframework.scheduling.annotation.EnableScheduling;
import org.springframework.web.servlet.LocaleResolver;
import org.springframework.web.servlet.config.annotation.InterceptorRegistry;
import org.springframework.web.servlet.i18n.LocaleChangeInterceptor;
import org.springframework.web.servlet.i18n.SessionLocaleResolver;

import com.o2osys.mng.common.storage.StorageProperties;

@SpringBootApplication
@EnableScheduling
@PropertySources(value = {
        @PropertySource("classpath:config.properties")
        ,@PropertySource("classpath:config-file.properties")
})
@EnableConfigurationProperties(StorageProperties.class)
public class ApiApplication {

    public static void main(String[] args) {
        SpringApplication.run(ApiApplication.class, args);
    }

    @Bean(name = "messageSource")
    public ReloadableResourceBundleMessageSource messageSource() {
        ReloadableResourceBundleMessageSource messageBundle = new ReloadableResourceBundleMessageSource();
        messageBundle.setBasename("classpath:messages/messages");
        messageBundle.setDefaultEncoding("UTF-8");
        messageBundle.setUseCodeAsDefaultMessage(true);
        return messageBundle;
    }

    @Bean(name = "localeResolver")
    public LocaleResolver sessionLocaleResolver() {
        //세션 기준으로 로케일을 설정 한다
        SessionLocaleResolver localeResolver = new SessionLocaleResolver();
        //쿠키 기준(세션이 끊겨도 브라우저에 설정된 쿠키 기준으로
        //CookieLocaleResolver localeResolver = new CookieLocalResolver();

        //최초 기본 로케일을 강제로 설정이 가능 하다.
        localeResolver.setDefaultLocale(new Locale("ko_KR"));
        return localeResolver;

    }

    @Bean
    public LocaleChangeInterceptor localeChangeInterceptor() {
        LocaleChangeInterceptor localeChangeInterceptor = new LocaleChangeInterceptor();
        //request로 넘어오는 language parameter를 받아서 locale로 설정 한다
        localeChangeInterceptor.setParamName("language");
        return localeChangeInterceptor;
    }

    public void addInterceptors(InterceptorRegistry registry) {
        //Interceptor를 추가 한다.
        registry.addInterceptor(localeChangeInterceptor());
    }

}