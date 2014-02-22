require 'mechanize'
require 'uri'
require 'axlsx'

require 'awesome_print'
require 'pry'

class Vcard
  attr_accessor :name, :postalcode, :city, :address, :website
  attr_reader :email, :phone

  def phone=(phone)
    @phone = phone.length > 1 ? phone.map { |p| p.inner_html.encode!('utf-8') }.join(', ') : phone.inner_html.encode!('utf-8')
  end

  def email=(email)
    @email = email.length > 1 ? email.map { |e| e.inner_html.encode!('utf-8') }.join(', ') : email.inner_html.encode!('utf-8')
  end
end

p         = Axlsx::Package.new
wb        = p.workbook
wb_header = %w( Név Cím Telefon Email Weboldal )

def next_link page
  link = page.link_with(text: 'következő oldal')
  link.href unless link.nil?
end

agent = Mechanize.new

page     = agent.get 'http://autoszerviz.helyek.eu/megye'
counties = page.links_with(:href => /^(\/megye)/)

puts 'starting to scrape county info'

counties.each do |county|
  puts "county: #{county.text}"
  vcards = []

  more_page = county.href
  loop do
    vcard_list  = agent.get more_page
    vcard_links = vcard_list.search("//div[@class='kereses_eredmeny']//a").map { |l| l.attr(:href) }.uniq
    vcard_links.each do |vcard_profile_link|
      vcard_page = agent.get vcard_profile_link

      puts "on page #{vcard_profile_link}"
      #sleep(1) # act as a friendly user, not as an agressive crawler bot

      vcard            = Vcard.new
      vcard.name       = vcard_page.search('h2.alcim2').inner_html.encode!('utf-8')
      vcard.postalcode = vcard_page.search('//span[@itemprop="postal-code"]').inner_html.encode!('utf-8')
      vcard.city       = vcard_page.search('//span[@itemprop="locality"]').inner_html.encode!('utf-8')
      vcard.address    = vcard_page.search('//span[@itemprop="street-address"]').inner_html.encode!('utf-8')
      vcard.phone      = vcard_page.search('//span[@itemprop="tel"]')
      vcard.email      = vcard_page.search('//span[@itemprop="email"]')
      vcard.website    = vcard_page.search('//a[@itemprop="url"]').inner_html.encode!('utf-8')

      vcards << vcard
    end

    more_page = next_link(vcard_list)
    break if more_page.nil?
  end

  # save vcards
  puts 'creating worksheet'

  wb.add_worksheet(:name => county.text) do |sheet|
    sheet.add_row wb_header
    vcards.each do |vcard|
      sheet.add_row [vcard.name, "#{vcard.city} #{vcard.postalcode}, #{vcard.address}", vcard.phone, vcard.email, vcard.website]
    end
  end
end

print 'dumping vcards to workbook...'
p.serialize('db.xlsx')
puts 'done'
